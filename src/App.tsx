/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, {
  useState,
  useEffect,
  useCallback,
  useMemo,
  useRef,
} from "react";
import {
  User,
  School,
  QrCode,
  UserPlus,
  GraduationCap,
  LayoutGrid,
  Settings,
  LogOut,
  ChevronRight,
  ChevronLeft,
  ChevronDown,
  ChevronUp,
  ClipboardList,
  FileBarChart,
  Table as TableIcon,
  Search,
  Plus,
  RefreshCcw,
  Printer,
  Download,
  Eye,
  EyeOff,
  Calendar,
  Clock,
  Trash2,
  Edit,
  Save,
  ArrowLeft,
  Upload,
  FileSpreadsheet,
  BarChart3,
  Info,
  CheckCircle2,
  XCircle,
  AlertTriangle,
  AlertCircle,
  Maximize2,
  CreditCard,
  Award,
  ExternalLink,
  ShieldCheck,
  Sparkles,
  Heart,
  ArrowUpCircle,
  BookOpen,
  X,
  FileText,
  CalendarRange,
  Bell,
  BellRing,
} from "lucide-react";
import { motion, AnimatePresence } from "framer-motion";
import { firestoreService } from "./services/firestoreService";
import * as XLSX from "xlsx";
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import * as htmlToImage from "html-to-image";
import { QRCodeSVG } from "qrcode.react";
import JSZip from "jszip";
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  Legend,
  AreaChart,
  Area,
} from "recharts";
import {
  UserSession,
  Student,
  Teacher,
  Classroom,
  Attendance,
  DashboardStats,
  Holiday,
  DaySetting,
  TeacherAttendance,
  AppConfig,
  TeachingSchedule,
} from "./types";
import { auth, db } from "./firebase";
import { signInAnonymously } from "firebase/auth";
import {
  doc,
  getDoc,
  setDoc,
  onSnapshot,
  collection,
  query,
  orderBy,
  where,
  getDocs,
  getCountFromServer,
  writeBatch,
} from "firebase/firestore";
import {
  handleFirestoreError,
  OperationType,
  onQuotaExceeded,
} from "./lib/firebaseUtils";

const formatIndoDate = (dateStr: string) => {
  if (!dateStr) return "-";
  try {
    // Handle YYYY-MM-DD to avoid timezone shifting
    if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
      const [y, m, d] = dateStr.split("-").map(Number);
      const date = new Date(y, m - 1, d);
      return new Intl.DateTimeFormat("id-ID", {
        weekday: "long",
        day: "numeric",
        month: "long",
        year: "numeric",
      }).format(date);
    }
    const d = new Date(dateStr);
    if (isNaN(d.getTime())) return dateStr;
    return new Intl.DateTimeFormat("id-ID", {
      weekday: "long",
      day: "numeric",
      month: "long",
      year: "numeric",
    }).format(d);
  } catch (e) {
    return dateStr;
  }
};

const formatIndoDateNoDay = (dateStr: string) => {
  if (!dateStr) return "-";
  try {
    if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
      const [y, m, d] = dateStr.split("-").map(Number);
      const date = new Date(y, m - 1, d);
      return new Intl.DateTimeFormat("id-ID", {
        day: "numeric",
        month: "long",
        year: "numeric",
      }).format(date);
    }
    const d = new Date(dateStr);
    if (isNaN(d.getTime())) return dateStr;
    return new Intl.DateTimeFormat("id-ID", {
      day: "numeric",
      month: "long",
      year: "numeric",
    }).format(d);
  } catch (e) {
    return dateStr;
  }
};

const getIndonesianDay = (dateStr: string) => {
  if (!dateStr) return "-";
  try {
    const [year, month, day] = dateStr.split("-").map(Number);
    const date = new Date(year, month - 1, day);
    const dayNames = [
      "Minggu",
      "Senin",
      "Selasa",
      "Rabu",
      "Kamis",
      "Jumat",
      "Sabtu",
    ];
    return dayNames[date.getDay()] || "-";
  } catch (e) {
    return "-";
  }
};

const getMonthYearText = (monthStr: string) => {
  if (!monthStr) return "-";
  try {
    const [y, m] = monthStr.split("-").map(Number);
    return new Date(y, m - 1, 1).toLocaleDateString("id-ID", {
      month: "long",
      year: "numeric",
    });
  } catch (e) {
    return monthStr;
  }
};

const countDaysInMonth = (
  year: number,
  month: number,
  dayName: string,
  holidays: Holiday[],
  upToDay?: number,
) => {
  const dayIndex = [
    "Minggu",
    "Senin",
    "Selasa",
    "Rabu",
    "Kamis",
    "Jumat",
    "Sabtu",
  ].indexOf(dayName);
  if (dayIndex === -1) return 0;

  let count = 0;
  let date = new Date(year, month, 1);
  const lastDay = upToDay || new Date(year, month + 1, 0).getDate();

  while (date.getMonth() === month && date.getDate() <= lastDay) {
    if (date.getDay() === dayIndex) {
      // Manual formatting YYYY-MM-DD in local time to match holiday data precisely
      const y = date.getFullYear();
      const m = String(date.getMonth() + 1).padStart(2, "0");
      const d = String(date.getDate()).padStart(2, "0");
      const dateStr = `${y}-${m}-${d}`;

      const isHoliday = (holidays || []).some((h) => h.tanggal === dateStr);
      if (!isHoliday) {
        count++;
      }
    }
    date.setDate(date.getDate() + 1);
  }
  return count;
};

const getEffectiveTargetSessionsForSchedule = (
  sch: TeachingSchedule,
  year: number,
  month: number,
  holidays: Holiday[],
  settings: DaySetting[],
  upToDay?: number,
) => {
  const dayIndex = [
    "Minggu",
    "Senin",
    "Selasa",
    "Rabu",
    "Kamis",
    "Jumat",
    "Sabtu",
  ].indexOf(sch.hari);
  if (dayIndex === -1) return 0;

  let totalSessions = 0;
  let date = new Date(year, month, 1);
  const lastDay = upToDay || new Date(year, month + 1, 0).getDate();

  while (date.getMonth() === month && date.getDate() <= lastDay) {
    if (date.getDay() === dayIndex) {
      const y = date.getFullYear();
      const m = String(date.getMonth() + 1).padStart(2, "0");
      const d = String(date.getDate()).padStart(2, "0");
      const dateStr = `${y}-${m}-${d}`;

      const isHoliday = (holidays || []).some((h) => h.tanggal === dateStr);
      if (!isHoliday) {
        const daySett = (settings || []).find((s) => s.hari === sch.hari);
        const limitJp = sch.hari === "Jumat" ? 6 : 8;
        const defaultJps = Array.from({ length: limitJp }, (_, i) => i + 1);
        let activeJps = defaultJps;
        if (daySett) {
          if (!daySett.targetDate || daySett.targetDate === dateStr) {
            activeJps = Array.isArray(daySett.activeJps)
              ? daySett.activeJps
              : defaultJps;
          }
        }

        const schJps = sch.jps || [];
        if (schJps.length > 0) {
          const activeSchJpCount = schJps.filter((j) =>
            activeJps.includes(j),
          ).length;
          totalSessions += activeSchJpCount;
        } else {
          const scale = activeJps.length / limitJp;
          totalSessions += Math.round(
            scale * (Number(sch.targetPertemuan) || 0),
          );
        }
      }
    }
    date.setDate(date.getDate() + 1);
  }
  return totalSessions;
};

const calculateTeacherMonthlyStats = (
  nip: string,
  bulan: string, // YYYY-MM
  kelasFilter: string,
  teachingSchedules: TeachingSchedule[],
  holidays: Holiday[],
  settings: DaySetting[],
  teacherAttendance: TeacherAttendance[],
) => {
  const [y, m] = bulan.split("-").map(Number);
  const now = new Date();
  const isCurrentMonth = now.getFullYear() === y && now.getMonth() + 1 === m;
  const isFutureMonth =
    y > now.getFullYear() ||
    (y === now.getFullYear() && m > now.getMonth() + 1);
  const lastDayOfMonth = new Date(y, m, 0).getDate();

  let totalTarget = 0;
  let hadir = 0;
  let sakit = 0;
  let izin = 0;
  let alfa = 0;
  let totalLambat = 0;

  const dayRecords = (teacherAttendance || []).filter(
    (ta) => ta.nip === nip && ta.tanggal.startsWith(bulan),
  );

  for (let d = 1; d <= lastDayOfMonth; d++) {
    const dateStr = `${y}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
    const currentDate = new Date(y, m - 1, d);
    const dayName = [
      "Minggu",
      "Senin",
      "Selasa",
      "Rabu",
      "Kamis",
      "Jumat",
      "Sabtu",
    ][currentDate.getDay()];

    if (dayName === "Minggu") continue;

    const isHoliday = (holidays || []).some((h) => h.tanggal === dateStr);
    if (isHoliday) continue;

    const daySett = (settings || []).find((s) => s.hari === dayName);
    const limitJp = dayName === "Jumat" ? 6 : 8;
    const defaultJps = Array.from({ length: limitJp }, (_, i) => i + 1);
    let activeJps = defaultJps;
    if (daySett) {
      if (!daySett.targetDate || daySett.targetDate === dateStr) {
        activeJps = Array.isArray(daySett.activeJps)
          ? daySett.activeJps
          : defaultJps;
      }
    }

    // Find schedules for this teacher today
    const schedules = (teachingSchedules || []).filter(
      (ts) =>
        ts.nip === nip &&
        ts.hari === dayName &&
        (kelasFilter ? ts.kelas === kelasFilter : true),
    );

    for (const sch of schedules) {
      const schJps = sch.jps || [];
      let activeSchJpCount = 0;
      if (schJps.length > 0) {
        activeSchJpCount = schJps.filter((j) => activeJps.includes(j)).length;
      } else {
        const scale = activeJps.length / limitJp;
        activeSchJpCount = Math.round(
          scale * (Number(sch.targetPertemuan) || 0),
        );
      }

      if (activeSchJpCount > 0) {
        totalTarget += activeSchJpCount;

        const isPastOrToday =
          !isFutureMonth && (!isCurrentMonth || d <= now.getDate());
        if (isPastOrToday) {
          // Check if scanned today
          const scans = dayRecords.filter(
            (ta) =>
              ta.tanggal === dateStr &&
              (!ta.status || ta.status === "Hadir") &&
              ta.kelas === sch.kelas,
          );

          const isSakit = dayRecords.some(
            (ta) =>
              ta.status === "Sakit" &&
              ta.tanggal === dateStr &&
              (ta.kelas === sch.kelas || ta.kelas === "-"),
          );
          const isIzin = dayRecords.some(
            (ta) =>
              ta.status === "Izin" &&
              ta.tanggal === dateStr &&
              (ta.kelas === sch.kelas || ta.kelas === "-"),
          );

          if (scans.length > 0) {
            hadir += activeSchJpCount;
            const lateVal = scans.reduce(
              (sum, s) => sum + (s.terlambat || 0),
              0,
            );
            totalLambat += lateVal;
          } else if (isSakit) {
            sakit += activeSchJpCount;
          } else if (isIzin) {
            izin += activeSchJpCount;
          } else {
            alfa += activeSchJpCount;
          }
        }
      }
    }
  }

  const perc = totalTarget > 0 ? Math.round((hadir / totalTarget) * 100) : 0;

  return {
    totalTarget,
    hadir,
    sakit,
    izin,
    alfa,
    totalLambat,
    perc,
  };
};

const getEffectiveDays = (
  year: number,
  month: number,
  holidays: Holiday[],
  settings: DaySetting[],
  upToDay?: number,
) => {
  const dayNames = [
    "Minggu",
    "Senin",
    "Selasa",
    "Rabu",
    "Kamis",
    "Jumat",
    "Sabtu",
  ];
  const days: string[] = [];
  const lastDay = upToDay || new Date(year, month + 1, 0).getDate();

  for (let d = 1; d <= lastDay; d++) {
    const date = new Date(year, month, d);
    const dayName = dayNames[date.getDay()];

    // 1. Skip Sunday
    if (dayName === "Minggu") continue;

    // 2. Skip Holidays
    const dateStr = `${year}-${String(month + 1).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
    if (holidays.some((h) => h.tanggal === dateStr)) continue;

    // 3. Skip if day setting is "Libur" (empty masuk or pulang)
    const setting = settings.find((s) => s.hari === dayName);
    if (setting && (!setting.masuk || !setting.pulang)) continue;

    days.push(dateStr);
  }
  return days;
};

const AttendanceChart = ({ data }: { data: any[] }) => (
  <div className="h-64 w-full">
    <ResponsiveContainer width="100%" height="100%" minWidth={0}>
      <AreaChart
        data={data}
        margin={{ top: 10, right: 30, left: 0, bottom: 0 }}
      >
        <defs>
          <linearGradient id="colorHadir" x1="0" y1="0" x2="0" y2="1">
            <stop offset="5%" stopColor="#16a34a" stopOpacity={0.8} />
            <stop offset="95%" stopColor="#16a34a" stopOpacity={0} />
          </linearGradient>
          <linearGradient id="colorAlfa" x1="0" y1="0" x2="0" y2="1">
            <stop offset="5%" stopColor="#dc2626" stopOpacity={0.8} />
            <stop offset="95%" stopColor="#dc2626" stopOpacity={0} />
          </linearGradient>
        </defs>
        <CartesianGrid
          strokeDasharray="3 3"
          vertical={false}
          stroke="#f1f5f9"
        />
        <XAxis dataKey="name" fontSize={10} axisLine={false} tickLine={false} />
        <YAxis fontSize={10} axisLine={false} tickLine={false} />
        <Tooltip
          contentStyle={{
            borderRadius: "12px",
            border: "none",
            boxShadow: "0 10px 15px -3px rgba(0, 0, 0, 0.1)",
          }}
          labelStyle={{ fontWeight: "bold" }}
        />
        <Legend iconType="circle" />
        <Area
          type="monotone"
          dataKey="hadir"
          stroke="#16a34a"
          fillOpacity={1}
          fill="url(#colorHadir)"
        />
        <Area
          type="monotone"
          dataKey="alfa"
          stroke="#dc2626"
          fillOpacity={1}
          fill="url(#colorAlfa)"
        />
      </AreaChart>
    </ResponsiveContainer>
  </div>
);

const formatMinutes = (minutes: number) => {
  if (minutes <= 0) return "Tepat Waktu";
  const h = Math.floor(minutes / 60);
  const m = minutes % 60;
  if (h > 0) return `${h} jam ${m} menit`;
  return `${m} menit`;
};

const getLocalISO = () => {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
};

export default function App() {
  const [students, setStudents] = useState<Student[]>([]);
  const [teachers, setTeachers] = useState<any[]>([]);
  const [classrooms, setClassrooms] = useState<Classroom[]>([]);
  const [attendance, setAttendance] = useState<Attendance[]>([]);
  const [teacherAttendance, setTeacherAttendance] = useState<
    TeacherAttendance[]
  >([]);
  const [settings, setSettings] = useState<DaySetting[]>([]);
  const [teachingSchedules, setTeachingSchedules] = useState<
    TeachingSchedule[]
  >([]);
  const [holidays, setHolidays] = useState<Holiday[]>([]);
  const [session, setSession] = useState<UserSession | null>(null);

  const [masterDataVersion, setMasterDataVersion] = useState(0);
  const refreshMasterData = useCallback(() => {
    setMasterDataVersion((prev) => prev + 1);
  }, []);

  const [offlineQueue, setOfflineQueue] = useState<any[]>(() => {
    try {
      return JSON.parse(localStorage.getItem("sigap_offline_attendance") || "[]");
    } catch {
      return [];
    }
  });

  const [offlineClassQueue, setOfflineClassQueue] = useState<any[]>(() => {
    try {
      return JSON.parse(localStorage.getItem("sigap_offline_class_attendance") || "[]");
    } catch {
      return [];
    }
  });

  const [syncingOffline, setSyncingOffline] = useState(false);

  const syncOfflineData = useCallback(async () => {
    if (syncingOffline) return;
    
    const attQueueStr = localStorage.getItem("sigap_offline_attendance");
    const classQueueStr = localStorage.getItem("sigap_offline_class_attendance");
    
    let attQueueObj: any[] = [];
    let classQueueObj: any[] = [];
    
    try {
      if (attQueueStr) attQueueObj = JSON.parse(attQueueStr);
    } catch (e) {
      console.error(e);
    }
    
    try {
      if (classQueueStr) classQueueObj = JSON.parse(classQueueStr);
    } catch (e) {
      console.error(e);
    }
    
    if (attQueueObj.length === 0 && classQueueObj.length === 0) return;
    
    if (!navigator.onLine) return; // Wait for actual internet
    
    setSyncingOffline(true);
    console.log(`[SIGAP] Syncing offline queues: ${attQueueObj.length} student, ${classQueueObj.length} coordinator logs.`);
    
    try {
      const batch = writeBatch(db);
      
      // Batch up student records
      attQueueObj.forEach((record) => {
        const docRef = doc(db, "attendance", record.id);
        batch.set(docRef, record);
      });
      
      // Batch up teacher records
      classQueueObj.forEach((record) => {
        const docRef = doc(db, "teacherAttendance", record.id);
        batch.set(docRef, record);
      });
      
      // Commit batch
      await batch.commit();
      console.log("[SIGAP] Batch commit of offline queue completed.");
      
      // Post-commit: update rekapSiswa summaries for each unique date of the student records
      const uniqueDates = Array.from(new Set(attQueueObj.map((r) => r.tanggal)));
      for (const tgl of uniqueDates) {
        await firestoreService.updateRekapSiswa(tgl);
      }
      
      // Clear Queues inside localStorage and React state
      localStorage.removeItem("sigap_offline_attendance");
      localStorage.removeItem("sigap_offline_class_attendance");
      setOfflineQueue([]);
      setOfflineClassQueue([]);
      
      triggerSuccess(
        "SINKRONISASI OFFLINE",
        `Sukses mengunggah ${attQueueObj.length} absensi siswa & ${classQueueObj.length} jurnal guru dari antrean offline.`
      );
    } catch (err) {
      console.error("Error committing offline sync:", err);
    } finally {
      setSyncingOffline(false);
    }
  }, [syncingOffline]);

  useEffect(() => {
    const handleOnline = () => {
      console.log("[SIGAP] Device is ONLINE. Triggering sync...");
      syncOfflineData();
    };
    window.addEventListener("online", handleOnline);
    return () => {
      window.removeEventListener("online", handleOnline);
    };
  }, [syncOfflineData]);

  useEffect(() => {
    const interval = setInterval(() => {
      if (navigator.onLine) {
        syncOfflineData();
      }
    }, 15000);
    return () => clearInterval(interval);
  }, [syncOfflineData]);

  const processOfflineScan = useCallback((scannedId: string, type: "Siswa" | "Kelas") => {
    const now = new Date();
    const dayMap = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
    const dayName = dayMap[now.getDay()];
    const tanggal = now.toISOString().split("T")[0];
    const jam = now.toLocaleTimeString("id-ID", {
      hour: "2-digit",
      minute: "2-digit",
      hour12: false,
    });

    if (dayName === "Minggu") {
      return { success: false, message: "Absen Gagal: Hari Minggu libur (Offline Mode)." };
    }

    // Check holidays from state safely
    const isHoliday = holidays.some((h) => h.tanggal === tanggal);
    if (isHoliday) {
      const hData = holidays.find((h) => h.tanggal === tanggal);
      return { success: false, message: `Absen Gagal: ${hData?.keterangan || "Hari Libur"} (Offline Mode).` };
    }

    if (type === "Siswa") {
      // Find student locally in the master students state!
      const student = students.find((s) => s.nisn === scannedId);
      if (!student) {
        return { success: false, message: "Absen Gagal: Siswa tidak terdaftar di memori offline." };
      }

      const id = `${scannedId}-${tanggal}`;

      // Check if already scanned in local state or offline queue
      const alreadyInQueue = offlineQueue.some((item) => item.id === id);
      const alreadyInAttendanceLocal = attendance.some((item) => item.id === id);
      if (alreadyInQueue || alreadyInAttendanceLocal) {
        return { success: false, message: `Absen Gagal: Anda sudah melakukan presensi hari ini (Offline Mode).` };
      }

      // Late calculations based on standard settings in memory
      let cutoffMasuk = "07:15";
      const sDay = settings.find((s) => s.hari === dayName);
      if (sDay?.masuk) cutoffMasuk = sDay.masuk;

      const currentMinutes = now.getHours() * 60 + now.getMinutes();
      const [hMasuk, mMasuk] = cutoffMasuk.split(":").map(Number);
      const masukMinutes = hMasuk * 60 + mMasuk;
      const late = currentMinutes - masukMinutes;
      const terlambat = late > 0 ? late : 0;

      const formatLateDescription = (min: number) => {
        if (min <= 0) return "Tepat Waktu";
        const h = Math.floor(min / 60);
        const m = min % 60;
        const hStr = h > 0 ? `${h} jam ` : "";
        const mStr = `${m} menit`;
        return `Terlambat ${hStr}${mStr}`;
      };

      const record = {
        id,
        nisn: scannedId,
        nama: student.nama,
        kelas: student.kelas,
        tanggal,
        jam,
        status: "Hadir",
        terlambat,
        keterangan: `${formatLateDescription(terlambat)} (Offline Scan)`,
      };

      // Add to offline queue
      const updatedQueue = [...offlineQueue, record];
      setOfflineQueue(updatedQueue);
      localStorage.setItem("sigap_offline_attendance", JSON.stringify(updatedQueue));

      // Append to local attendance state immediately
      setAttendance((prev) => [record as any, ...prev]);

      return {
        success: true,
        status: "Hadir",
        message: `[OFFLINE] Berhasil Absen: ${student.nama}. Disimpan di antrean lokal.`,
      };
    } else {
      // Teacher Class Scan
      const className = scannedId;
      const nip = session?.uid || "";
      const id = `${nip}-${className}-${now.getTime()}`;

      // Lookup teaching schedules in memory to find subject
      const schedule = teachingSchedules.find(
        (ts) => ts.nip === nip && ts.kelas === className
      );
      const mapel = schedule?.mapel || "-";

      const record = {
        id,
        nip,
        nama: session?.name || "Guru",
        kelas: className,
        tanggal,
        jam,
        terlambat: 0,
        mapel,
        keterangan: "Kelas (Offline Scan)",
      };

      const updatedQueue = [...offlineClassQueue, record];
      setOfflineClassQueue(updatedQueue);
      localStorage.setItem("sigap_offline_class_attendance", JSON.stringify(updatedQueue));

      // Append to local teacherAttendance state so they see it in the list!
      setTeacherAttendance((prev) => [record as any, ...prev]);

      return {
        success: true,
        message: `[OFFLINE] Berhasil Scan Kelas: ${className}. Disimpan di antrean lokal.`,
      };
    }
  }, [students, settings, holidays, offlineQueue, offlineClassQueue, attendance, teachingSchedules, session]);
  const [savedDaysStatus, setSavedDaysStatus] = useState<{
    [hari: string]: "idle" | "saving" | "saved";
  }>({});
  const [localActiveJps, setLocalActiveJps] = useState<{
    [hari: string]: number[];
  }>({});
  const [localJpTimes, setLocalJpTimes] = useState<{
    [hari: string]: { [jp: number]: { start: string; end: string } };
  }>({});
  const [expandedJpDays, setExpandedJpDays] = useState<{
    [hari: string]: boolean;
  }>({});
  const [currentTime, setCurrentTime] = useState("");
  const [expandedKamadTeacherDetails, setExpandedKamadTeacherDetails] =
    useState<{ [nip: string]: boolean }>({});
  const [expandedJadwal, setExpandedJadwal] = useState<string[]>([]);

  // States for Input Kehadiran Guru manual
  const [manT_nip, setManT_nip] = useState("");
  const [manT_status, setManT_status] = useState<"Sakit" | "Izin">("Sakit");
  const [manT_mode, setManT_mode] = useState<"single" | "range">("single");
  const [manT_date, setManT_date] = useState("");
  const [manT_startDate, setManT_startDate] = useState("");
  const [manT_endDate, setManT_endDate] = useState("");
  const [manT_keterangan, setManT_keterangan] = useState("");
  const [manT_saving, setManT_saving] = useState(false);
  const [manT_searchTerm, setManT_searchTerm] = useState("");

  useEffect(() => {
    const updateTime = () => {
      const now = new Date();
      const hh = String(now.getHours()).padStart(2, "0");
      const mm = String(now.getMinutes()).padStart(2, "0");
      setCurrentTime(`${hh}:${mm}`);
    };
    updateTime();
    const timer = setInterval(updateTime, 1000);
    return () => clearInterval(timer);
  }, []);

  const [localReasonInactive, setLocalReasonInactive] = useState<{
    [hari: string]: string;
  }>({});
  const [localMasuk, setLocalMasuk] = useState<{ [hari: string]: string }>({});
  const [localPulang, setLocalPulang] = useState<{ [hari: string]: string }>(
    {},
  );
  const [appConfig, setAppConfig] = useState<AppConfig>({});

  const [searchTermGuru, setSearchTermGuru] = useState("");
  const [searchTermSiswa, setSearchTermSiswa] = useState("");
  const [searchTermKelas, setSearchTermKelas] = useState("");
  const [searchTermJadwal, setSearchTermJadwal] = useState("");
  const [refreshing, setRefreshing] = useState(false);
  const [studentProfileClassFilter, setStudentProfileClassFilter] =
    useState("");

  // Parent Notification Center States for Student Dashboard
  const [parentNotifEnabled, setParentNotifEnabled] = useState(() => {
    return localStorage.getItem("parent_notif_enabled") === "true";
  });
  const [notifPermission, setNotifPermission] =
    useState<NotificationPermission>(() => {
      if (typeof window !== "undefined" && "Notification" in window) {
        try {
          return Notification.permission as NotificationPermission;
        } catch (e) {
          console.warn("Could not read Notification.permission:", e);
          return "default";
        }
      }
      return "default";
    });
  const [isInIframe, setIsInIframe] = useState(false);

  useEffect(() => {
    if (typeof window !== "undefined") {
      setIsInIframe(window.self !== window.top);

      // Update permission status dynamically when tab gains focus
      const updatePermission = () => {
        if ("Notification" in window) {
          try {
            setNotifPermission(
              Notification.permission as NotificationPermission,
            );
          } catch (e) {
            console.warn(e);
          }
        }
      };

      window.addEventListener("focus", updatePermission);
      return () => {
        window.removeEventListener("focus", updatePermission);
      };
    }
  }, []);

  const isFirstLoadRef = useRef(true);
  const previousAttendanceRef = useRef<
    Record<string, { status: string; jam: string; jamPulang?: string }>
  >({});

  const triggerParentNotification = useCallback(
    (title: string, body: string) => {
      // 1. Browser Native push notification
      if (
        typeof window !== "undefined" &&
        "Notification" in window &&
        Notification.permission === "granted"
      ) {
        try {
          new Notification(title, {
            body,
            icon: appConfig.logoUrl || "/favicon.ico",
          });
        } catch (e) {
          console.error("Failed to show native browser notification", e);
        }
      }

      // 2. Beautiful app success toast as active indicator
      triggerSuccess(title.toUpperCase(), body);
    },
    [appConfig.logoUrl],
  );

  // Monitor student attendance records live for Parent Notifications
  useEffect(() => {
    // Only monitor if logged in as Siswa
    if (session?.role !== "Siswa") {
      isFirstLoadRef.current = true;
      return;
    }

    const todayStr = new Date().toLocaleDateString("en-CA"); // YYYY-MM-DD local
    const studentAttendance = attendance.filter((a) => a.nisn === session.uid);

    if (isFirstLoadRef.current) {
      // Initialize map of current attendance records
      const initialMap: Record<
        string,
        { status: string; jam: string; jamPulang?: string }
      > = {};
      studentAttendance.forEach((a) => {
        initialMap[a.id] = {
          status: a.status || "",
          jam: a.jam || "",
          jamPulang: a.jamPulang || "",
        };
      });
      previousAttendanceRef.current = initialMap;
      isFirstLoadRef.current = false;
      return;
    }

    if (!parentNotifEnabled) return;

    // Check for changes/additions in today's attendance records
    studentAttendance.forEach((a) => {
      // Only trigger if date matches today local
      if (a.tanggal !== todayStr) {
        previousAttendanceRef.current[a.id] = {
          status: a.status || "",
          jam: a.jam || "",
          jamPulang: a.jamPulang || "",
        };
        return;
      }

      const prev = previousAttendanceRef.current[a.id];

      // New attendance record created for today
      if (!prev) {
        if (a.jam) {
          triggerParentNotification(
            "📍 Presensi Masuk Siswa",
            `Anak Anda (${a.nama}) telah terdeteksi hadir di madrasah pada pukul ${a.jam}. Status: ${a.status || "Hadir"}.`,
          );
        }
      }
      // Update on today's record (e.g. status status or check-in speed)
      else {
        if (a.jam && prev.jam !== a.jam) {
          triggerParentNotification(
            "📍 Update Kehadiran",
            `Update: Anak Anda (${a.nama}) terpantau masuk pada pukul ${a.jam}. Status: ${a.status || "Hadir"}.`,
          );
        }
        if (a.jamPulang && !prev.jamPulang) {
          triggerParentNotification(
            "🚪 Presensi Pulang Siswa",
            `Anak Anda (${a.nama}) telah melakukan presensi PULANG dari madrasah pada pukul ${a.jamPulang}.`,
          );
        }
      }

      // Sync ref values
      previousAttendanceRef.current[a.id] = {
        status: a.status || "",
        jam: a.jam || "",
        jamPulang: a.jamPulang || "",
      };
    });
  }, [attendance, session, parentNotifEnabled, triggerParentNotification]);

  const [activePanel, setActivePanel] = useState("home");

  const [analysisClass, setAnalysisClass] = useState("");
  const [analysisMonth, setAnalysisMonth] = useState(getLocalISO().slice(0, 7));
  const [appLogoInput, setAppLogoInput] = useState("");

  // Force analysisClass for Wali Kelas
  useEffect(() => {
    if (
      activePanel === "analisis" &&
      session?.role === "Guru" &&
      session?.isWali &&
      session?.kelas
    ) {
      setAnalysisClass(session.kelas);
    }
  }, [activePanel, session]);

  const effectiveDaysThisMonth = useMemo(() => {
    const now = new Date();
    return getEffectiveDays(
      now.getFullYear(),
      now.getMonth(),
      holidays,
      settings,
      now.getDate(),
    );
  }, [holidays, settings]);

  const effectiveDaysForAnalysisMonth = useMemo(() => {
    if (!analysisMonth) return [];
    try {
      const [y, m] = analysisMonth.split("-").map(Number);
      const now = new Date();
      const currentMonthPrefix = now.toISOString().slice(0, 7);

      let limit;
      if (analysisMonth < currentMonthPrefix) {
        limit = new Date(y, m, 0).getDate();
      } else if (analysisMonth === currentMonthPrefix) {
        limit = now.getDate();
      } else {
        limit = 0;
      }
      return getEffectiveDays(y, m - 1, holidays, settings, limit);
    } catch (e) {
      return [];
    }
  }, [analysisMonth, holidays, settings]);

  const DEFAULT_SUBJECTS = useMemo(
    () => [
      "Fiqih",
      "Alqur'an Hadis",
      "Akida Akhlak",
      "SKI",
      "Bhs. Indonesia",
      "Bhs. Arab",
      "Bhs. Inggris",
      "Matematika",
      "IPA",
      "IPS",
      "Penjas",
      "PPKN",
      "Informatika",
      "Koding K.A",
      "Mulok",
    ],
    [],
  );

  const [newSubject, setNewSubject] = useState("");
  const activeSubjects = useMemo(() => {
    return Array.isArray(appConfig.subjects)
      ? appConfig.subjects
      : DEFAULT_SUBJECTS;
  }, [appConfig.subjects, DEFAULT_SUBJECTS]);

  const [selectedDetailRecap, setSelectedDetailRecap] = useState<{
    kelas: string;
    tanggal: string;
    mapel: string;
    namaGuru: string;
  } | null>(null);

  const [showMonthlyMapelAttendance, setShowMonthlyMapelAttendance] = useState(false);

  const [loading, setLoading] = useState(false);
  const [loaderText, setLoaderText] = useState("");
  const [showMassDownloadConfirm, setShowMassDownloadConfirm] = useState(false);
  const [massDownloadSiswa, setMassDownloadSiswa] = useState<Student[]>([]);
  const [firebaseConnected, setFirebaseConnected] = useState(false);
  const [quotaExceeded, setQuotaExceeded] = useState(false);
  const [dismissQuotaNotice, setDismissQuotaNotice] = useState(false);
  const [anonymousAuthError, setAnonymousAuthError] = useState<string | null>(null);

  const [globalSiswaCount, setGlobalSiswaCount] = useState<number>(0);
  const [globalGuruCount, setGlobalGuruCount] = useState<number>(0);
  const [globalSiswaL, setGlobalSiswaL] = useState<number>(0);
  const [globalSiswaP, setGlobalSiswaP] = useState<number>(0);

  // Fetch school-wide counts for fallbacks in user accounts (homeroom teachers / general teachers)
  useEffect(() => {
    if (!firebaseConnected || !session) {
      setGlobalSiswaCount(0);
      setGlobalGuruCount(0);
      setGlobalSiswaL(0);
      setGlobalSiswaP(0);
      return;
    }

    const fetchGlobalCounts = async () => {
      try {
        const [siswaSnap, guruSnap, siswaLSnap] = await Promise.all([
          getCountFromServer(query(collection(db, "students"))),
          getCountFromServer(query(collection(db, "teachers"))),
          getCountFromServer(
            query(
              collection(db, "students"),
              where("jenisKelamin", "in", [
                "L",
                "Laki-Laki",
                "Laki-laki",
                "Laki-Laki ",
              ]),
            ),
          ),
        ]);

        const siswaCount = siswaSnap.data().count;
        const guruCount = guruSnap.data().count;
        const siswaL = siswaLSnap.data().count;
        const siswaP = Math.max(0, siswaCount - siswaL);

        setGlobalSiswaCount(siswaCount);
        setGlobalGuruCount(guruCount);
        setGlobalSiswaL(siswaL);
        setGlobalSiswaP(siswaP);
      } catch (err) {
        console.error("Error fetching global school counts:", err);
      }
    };

    fetchGlobalCounts();
  }, [firebaseConnected, session]);

  const [filterClassAbsensi, setFilterClassAbsensi] = useState("");
  const [filterGuruAbsensiClass, setFilterGuruAbsensiClass] = useState("");
  const [filterTanggalAbsensi, setFilterTanggalAbsensi] = useState("");
  const [filterTanggalAbsensiGuru, setFilterTanggalAbsensiGuru] = useState("");
  const [filterNamaAbsensi, setFilterNamaAbsensi] = useState("");
  const [filterAdminSiswaClass, setFilterAdminSiswaClass] = useState("");
  const [filterJadwalClass, setFilterJadwalClass] = useState("");
  const [siswaDashboardDate, setSiswaDashboardDate] = useState(
    new Date().toISOString().split("T")[0],
  );
  const [waliFilterTanggal, setWaliFilterTanggal] = useState(
    new Date().toISOString().split("T")[0],
  );
  const [waliFilterNama, setWaliFilterNama] = useState("");
  const [waliFilterStatus, setWaliFilterStatus] = useState("");
  const [waliPagination, setWaliPagination] = useState(0);

  const [pagination, setPagination] = useState({
    guru: 0,
    siswa: 0,
    kelas: 0,
    jadwal: 0,
    absensiSiswa: 0,
    absensiGuru: 0,
    rekapMapel: 0,
    rekapSiswa: 0,
    rekapGuru: 0,
  });
  const [pageSize, setPageSize] = useState(10);
  const [rekapDailyList, setRekapDailyList] = useState<any[]>([]);

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
    nip: "",
    bulan: getLocalISO().slice(0, 7),
    mapel: "",
    kelas: "",
  });

  useEffect(() => {
    setPagination((p) => ({ ...p, rekapMapel: 0 }));
  }, [
    rekapMapelFilter.nip,
    rekapMapelFilter.bulan,
    rekapMapelFilter.mapel,
    rekapMapelFilter.kelas,
    pageSize,
  ]);

  const [confirmModal, setConfirmModal] = useState<{
    show: boolean;
    title: string;
    message: string;
    entityName?: string;
    onConfirm: () => void;
  } | null>(null);
  const [resetAbsensiTarget, setResetAbsensiTarget] = useState<
    "Siswa" | "Guru" | null
  >(null);
  const [resetAbsensiBulan, setResetAbsensiBulan] = useState(
    getLocalISO().slice(0, 7),
  );

  useEffect(() => {
    onQuotaExceeded((isExceeded) => {
      setQuotaExceeded(isExceeded);
    });

    const initApp = async () => {
      setLoading(true);
      try {
        // Enforce anonymous auth on boot for role-based access validation
        try {
          await signInAnonymously(auth);
        } catch (authErr: any) {
          console.warn("[Auth] Anonymous Auth failed:", authErr);
          setAnonymousAuthError(authErr?.message || String(authErr));
        }

        // Automated data migration: migrate any legacy teachers to isolated subcollection
        try {
          const q = query(collection(db, "teachers"));
          const snapshot = await getDocs(q);
          for (const docSnap of snapshot.docs) {
            const data = docSnap.data();
            if (data.pass) {
              console.log(`[Migration] Safely migrating password for ${docSnap.id} to isolated subcollection`);
              await setDoc(doc(db, "teachers", docSnap.id, "private", "password"), {
                pass: data.pass
              });
              const { pass, ...publicData } = data;
              await setDoc(doc(db, "teachers", docSnap.id), publicData);
            }
          }
        } catch (migErr) {
          console.log("[Migration] Skipping migration:", migErr);
        }

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

    const unsubAppConfig = onSnapshot(
      doc(db, "appConfig", "general"),
      (snap) => {
        if (snap.exists()) {
          const data = snap.data() as AppConfig;
          setAppConfig(data);
        }
      },
      (error) =>
        handleFirestoreError(error, OperationType.GET, "appConfig/general"),
    );

    return () => {
      unsubAppConfig();
    };
  }, [firebaseConnected]);

  // Load transaction data ONLY when a user session is active (logged in)
  useEffect(() => {
    if (!firebaseConnected || !session) {
      // Clear data states when logged out to avoid stale/previous user data and stop memory/network consumption
      setAttendance([]);
      setTeacherAttendance([]);
      setSettings([]);
      setHolidays([]);
      setRekapDailyList([]);
      return;
    }

    // Real-time synchronization for critical data (restricted to last 30 days)
    const thirtyDaysAgo = new Date();
    thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
    const dateLimit = thirtyDaysAgo.toISOString().split("T")[0];

    const role = session.role;
    const uid = session.uid;
    const kelasWali = session.kelas;
    const isWali = session.isWali;
    const jabatan = session.jabatan || "";

    // Deterministic state check for school leadership
    const isLeadership =
      role === "Guru" && (jabatan === "Kamad" || jabatan === "Wakamad");

    let unsubOwnProfile = () => {};
    let unsubAttendance = () => {};
    let unsubTeacherAttendance = () => {};
    let unsubSettings = () => {};
    let unsubHolidays = () => {};
    let unsubRekapSiswa = () => {};

    // 1. Settings, Holidays & RekapSiswa are lightweight config models, fetch safely
    unsubSettings = onSnapshot(
      collection(db, "settings"),
      (snap) => {
        setSettings(snap.docs.map((d) => d.data() as DaySetting));
      },
      (error) => handleFirestoreError(error, OperationType.GET, "settings"),
    );

    unsubHolidays = onSnapshot(
      collection(db, "holidays"),
      (snap) => {
        const data = snap.docs.map((d) => d.data() as Holiday);
        setHolidays(data.sort((a, b) => b.tanggal.localeCompare(a.tanggal)));
      },
      (error) => handleFirestoreError(error, OperationType.GET, "holidays"),
    );

    unsubRekapSiswa = onSnapshot(
      query(
        collection(db, "rekapSiswa"),
        where("tanggal", ">=", dateLimit),
        orderBy("tanggal", "desc"),
      ),
      (snap) => {
        setRekapDailyList(snap.docs.map((d) => d.data()));
      },
      (error) => handleFirestoreError(error, OperationType.GET, "rekapSiswa"),
    );

    // Real-time synchronization of teacher profile to react instantly to admin role/position/jabatan changes
    if (role === "Guru") {
      unsubOwnProfile = onSnapshot(
        doc(db, "teachers", uid),
        (snap) => {
          if (snap.exists()) {
            const tData = snap.data() as Teacher;
            setSession((prev) => {
              if (!prev) return prev;
              const updatedIsWali = !!tData.kelas;
              const hasDiff =
                prev.name !== tData.nama ||
                prev.kelas !== tData.kelas ||
                prev.isWali !== updatedIsWali ||
                prev.jabatan !== tData.jabatan;

              if (hasDiff) {
                console.log(
                  "[SIGAP] Real-time session update detected:",
                  tData,
                );
                return {
                  ...prev,
                  name: tData.nama,
                  kelas: tData.kelas,
                  isWali: updatedIsWali,
                  jabatan: tData.jabatan,
                };
              }
              return prev;
            });
          }
        },
        (error) => console.error("Error syncing own profile", error),
      );
    }

    // 2. Performance & Quota Optimization dynamically applied for logs
    if (role === "Siswa") {
      unsubAttendance = onSnapshot(
        query(collection(db, "attendance"), where("nisn", "==", uid)),
        (snap) => {
          setAttendance(snap.docs.map((d) => d.data() as Attendance));
        },
        (error) => handleFirestoreError(error, OperationType.GET, "attendance"),
      );
    } else if (role === "Guru") {
      if (isLeadership) {
        unsubAttendance = onSnapshot(
          query(
            collection(db, "attendance"),
            where("tanggal", ">=", dateLimit),
            orderBy("tanggal", "desc"),
          ),
          (snap) => {
            setAttendance(snap.docs.map((d) => d.data() as Attendance));
          },
          (error) =>
            handleFirestoreError(error, OperationType.GET, "attendance"),
        );

        unsubTeacherAttendance = onSnapshot(
          query(
            collection(db, "teacherAttendance"),
            where("tanggal", ">=", dateLimit),
            orderBy("tanggal", "desc"),
          ),
          (snap) => {
            setTeacherAttendance(
              snap.docs.map((d) => d.data() as TeacherAttendance),
            );
          },
          (error) =>
            handleFirestoreError(error, OperationType.GET, "teacherAttendance"),
        );
      } else if (isWali && kelasWali) {
        unsubAttendance = onSnapshot(
          query(
            collection(db, "attendance"),
            where("kelas", "==", kelasWali),
            where("tanggal", ">=", dateLimit),
          ),
          (snap) => {
            setAttendance(snap.docs.map((d) => d.data() as Attendance));
          },
          (error) =>
            handleFirestoreError(error, OperationType.GET, "attendance"),
        );

        unsubTeacherAttendance = onSnapshot(
          query(
            collection(db, "teacherAttendance"),
            where("tanggal", ">=", dateLimit),
          ),
          (snap) => {
            setTeacherAttendance(
              snap.docs.map((d) => d.data() as TeacherAttendance),
            );
          },
          (error) =>
            handleFirestoreError(error, OperationType.GET, "teacherAttendance"),
        );
      } else {
        unsubTeacherAttendance = onSnapshot(
          query(
            collection(db, "teacherAttendance"),
            where("nip", "==", uid),
            where("tanggal", ">=", dateLimit),
          ),
          (snap) => {
            setTeacherAttendance(
              snap.docs.map((d) => d.data() as TeacherAttendance),
            );
          },
          (error) =>
            handleFirestoreError(error, OperationType.GET, "teacherAttendance"),
        );
      }
    } else if (role === "Admin") {
      unsubAttendance = onSnapshot(
        query(
          collection(db, "attendance"),
          where("tanggal", ">=", dateLimit),
          orderBy("tanggal", "desc"),
        ),
        (snap) => {
          setAttendance(snap.docs.map((d) => d.data() as Attendance));
        },
        (error) => handleFirestoreError(error, OperationType.GET, "attendance"),
      );

      unsubTeacherAttendance = onSnapshot(
        query(
          collection(db, "teacherAttendance"),
          where("tanggal", ">=", dateLimit),
          orderBy("tanggal", "desc"),
        ),
        (snap) => {
          setTeacherAttendance(
            snap.docs.map((d) => d.data() as TeacherAttendance),
          );
        },
        (error) =>
          handleFirestoreError(error, OperationType.GET, "teacherAttendance"),
      );
    }

    return () => {
      unsubOwnProfile();
      unsubAttendance();
      unsubTeacherAttendance();
      unsubSettings();
      unsubHolidays();
      unsubRekapSiswa();
    };
  }, [
    firebaseConnected,
    session?.uid,
    session?.role,
    session?.kelas,
    session?.isWali,
    session?.jabatan,
  ]);

  // One-time Fetch for master data (students, teachers, classrooms, schedules)
  useEffect(() => {
    if (!firebaseConnected || !session) {
      setStudents([]);
      setTeachers([]);
      setClassrooms([]);
      setTeachingSchedules([]);
      return;
    }

    const loadMasterData = async () => {
      const role = session.role;
      const uid = session.uid;
      const kelasWali = session.kelas;
      const isWali = session.isWali;
      const jabatan = session.jabatan || "";
      const isLeadership =
        role === "Guru" && (jabatan === "Kamad" || jabatan === "Wakamad");

      try {
        if (role === "Siswa") {
          const snap = await getDoc(doc(db, "students", uid));
          if (snap.exists()) {
            setStudents([{ nisn: snap.id, ...snap.data() } as Student]);
          } else {
            setStudents([]);
          }
          setTeachers([]);
          setClassrooms([]);
          setTeachingSchedules([]);
        } else if (role === "Guru") {
          if (isLeadership) {
            const [studentsSnap, teachersSnap, classroomsSnap, schedulesSnap] = await Promise.all([
              getDocs(collection(db, "students")),
              getDocs(collection(db, "teachers")),
              getDocs(collection(db, "classrooms")),
              getDocs(collection(db, "teachingSchedules")),
            ]);

            const sData = studentsSnap.docs.map(
              (d) => ({ nisn: d.id, ...d.data() }) as Student,
            );
            setStudents(sData.sort((a, b) => a.nama.localeCompare(b.nama)));

            const tData = teachersSnap.docs.map(
              (d) => ({ nip: d.id, ...d.data() }) as Teacher,
            );
            setTeachers(tData.sort((a, b) => a.nama.localeCompare(b.nama)));

            const cData = classroomsSnap.docs.map((d) => d.data() as Classroom);
            setClassrooms(cData.sort((a, b) => a.nama.localeCompare(b.nama)));

            setTeachingSchedules(
              schedulesSnap.docs.map((d) => d.data() as TeachingSchedule),
            );
          } else if (isWali && kelasWali) {
            const [studentsSnap, teachersSnap, classroomsSnap, schedulesSnap] = await Promise.all([
              getDocs(query(collection(db, "students"), where("kelas", "==", kelasWali))),
              getDocs(collection(db, "teachers")),
              getDocs(collection(db, "classrooms")),
              getDocs(collection(db, "teachingSchedules")),
            ]);

            const sData = studentsSnap.docs.map(
              (d) => ({ nisn: d.id, ...d.data() }) as Student,
            );
            setStudents(sData.sort((a, b) => a.nama.localeCompare(b.nama)));

            const tData = teachersSnap.docs.map(
              (d) => ({ nip: d.id, ...d.data() }) as Teacher,
            );
            setTeachers(tData.sort((a, b) => a.nama.localeCompare(b.nama)));

            const cData = classroomsSnap.docs.map((d) => d.data() as Classroom);
            setClassrooms(cData.sort((a, b) => a.nama.localeCompare(b.nama)));

            setTeachingSchedules(
              schedulesSnap.docs.map((d) => d.data() as TeachingSchedule),
            );
          } else {
            const [teacherSnap, schedulesSnap] = await Promise.all([
              getDoc(doc(db, "teachers", uid)),
              getDocs(query(collection(db, "teachingSchedules"), where("nip", "==", uid))),
            ]);

            if (teacherSnap.exists()) {
              setTeachers([{ nip: teacherSnap.id, ...teacherSnap.data() } as Teacher]);
            } else {
              setTeachers([]);
            }
            setTeachingSchedules(
              schedulesSnap.docs.map((d) => d.data() as TeachingSchedule),
            );
            setStudents([]);
            setClassrooms([]);
          }
        } else if (role === "Admin") {
          const [studentsSnap, teachersSnap, classroomsSnap, schedulesSnap] = await Promise.all([
            getDocs(collection(db, "students")),
            getDocs(collection(db, "teachers")),
            getDocs(collection(db, "classrooms")),
            getDocs(collection(db, "teachingSchedules")),
          ]);

          const sData = studentsSnap.docs.map(
            (d) => ({ nisn: d.id, ...d.data() }) as Student,
          );
          setStudents(sData.sort((a, b) => a.nama.localeCompare(b.nama)));

          const tData = teachersSnap.docs.map(
            (d) => ({ nip: d.id, ...d.data() }) as Teacher,
          );
          setTeachers(tData.sort((a, b) => a.nama.localeCompare(b.nama)));

          const cData = classroomsSnap.docs.map((d) => d.data() as Classroom);
          setClassrooms(cData.sort((a, b) => a.nama.localeCompare(b.nama)));

          setTeachingSchedules(
            schedulesSnap.docs.map((d) => d.data() as TeachingSchedule),
          );
        }
      } catch (err) {
        console.error("Error loading master data: ", err);
      }
    };

    loadMasterData();
  }, [firebaseConnected, session?.uid, session?.role, session?.kelas, session?.isWali, masterDataVersion]);

  const [selectedTeacherNipForCapaian, setSelectedTeacherNipForCapaian] =
    useState<string | null>(null);

  const stats: DashboardStats & {
    siswaL?: number;
    siswaP?: number;
    sakitCount?: number;
    izinCount?: number;
    alfaCount?: number;
  } = useMemo(() => {
    const now = new Date();
    const today = currentDate; // Use the state-synced date
    const dayMap = [
      "Minggu",
      "Senin",
      "Selasa",
      "Rabu",
      "Kamis",
      "Jumat",
      "Sabtu",
    ];
    const dayName = dayMap[now.getDay()];

    const isHoliday =
      holidays.some((h) => h.tanggal === today) || dayName === "Minggu";

    const todayRekap = rekapDailyList.find((r) => r.tanggal === today);
    const latestRekap = rekapDailyList.length > 0 ? rekapDailyList[0] : null;

    let siswaCount =
      todayRekap?.siswaCount ??
      latestRekap?.siswaCount ??
      (globalSiswaCount > 0 ? globalSiswaCount : students.length);
    let guruCount =
      todayRekap?.guruCount ??
      latestRekap?.guruCount ??
      (globalGuruCount > 0 ? globalGuruCount : teachers.length);
    let hadir = todayRekap?.hadirCount ?? 0;
    let lambat = todayRekap?.terlambatCount ?? 0;
    let sakit = todayRekap?.sakitCount ?? 0;
    let izin = todayRekap?.izinCount ?? 0;
    let alfaTotal = todayRekap?.alfaCount ?? 0;
    let siswaL =
      todayRekap?.siswaL ??
      latestRekap?.siswaL ??
      (globalSiswaL > 0
        ? globalSiswaL
        : students.filter(
            (s) => s.jenisKelamin === "L" || s.jenisKelamin === "Laki-Laki",
          ).length);
    let siswaP =
      todayRekap?.siswaP ??
      latestRekap?.siswaP ??
      (globalSiswaP > 0
        ? globalSiswaP
        : students.filter(
            (s) => s.jenisKelamin === "P" || s.jenisKelamin === "Perempuan",
          ).length);

    // Use live fallback ONLY for Admin and Leadership when rekap is completely unavailable and they actually have full lists loaded (> 50 students)
    const isLeadershipOrAdmin =
      session?.role === "Admin" ||
      (session?.role === "Guru" &&
        (session?.jabatan === "Kamad" || session?.jabatan === "Wakamad"));
    if (!todayRekap) {
      if (isLeadershipOrAdmin && students.length > 50) {
        const todayAttendance = attendance.filter((a) => a.tanggal === today);
        hadir = todayAttendance.filter((a) => a.status === "Hadir").length;
        lambat = todayAttendance.filter((a) => a.terlambat > 0).length;
        sakit = todayAttendance.filter((a) => a.status === "Sakit").length;
        izin = todayAttendance.filter((a) => a.status === "Izin").length;
        const recordedAlfa = todayAttendance.filter(
          (a) => a.status === "Alfa",
        ).length;
        const attendedNisns = new Set(todayAttendance.map((a) => a.nisn));
        const notYetCheckedIn = students.length - attendedNisns.size;
        alfaTotal = isHoliday ? 0 : recordedAlfa + notYetCheckedIn;
      } else {
        // For general teachers / Wali Kelas, if today's summary is not in DB yet, attendance stats are 0,
        // and Alfa defaults to the total school-wide student count (every student is technically not checked-in yet)
        alfaTotal = isHoliday ? 0 : siswaCount;
      }
    }

    // Chart data (last 7 days)
    const chartData = [];
    for (let i = 6; i >= 0; i--) {
      const d = new Date();
      d.setDate(d.getDate() - i);
      const dateStr = d.toISOString().split("T")[0];
      const dayRekap = rekapDailyList.find((r) => r.tanggal === dateStr);

      if (dayRekap) {
        chartData.push({
          name: d.toLocaleDateString("id-ID", { weekday: "short" }),
          hadir: dayRekap.hadirCount ?? 0,
          alfa: dayRekap.alfaCount ?? 0,
        });
      } else {
        // Fallback to live attendance if available
        const dayData = attendance.filter((a) => a.tanggal === dateStr);
        const dayHadir = dayData.filter((a) => a.status === "Hadir").length;
        const dDayMap = [
          "Minggu",
          "Senin",
          "Selasa",
          "Rabu",
          "Kamis",
          "Jumat",
          "Sabtu",
        ];
        const dDayName = dDayMap[d.getDay()];
        const dIsHoliday =
          holidays.some((h) => h.tanggal === dateStr) || dDayName === "Minggu";
        const dNotYet =
          students.length > 0
            ? students.length - new Set(dayData.map((a) => a.nisn)).size
            : 0;
        const dRecordedAlfa = dayData.filter((a) => a.status === "Alfa").length;
        const dAlfaTotal = dIsHoliday ? 0 : dNotYet + dRecordedAlfa;

        chartData.push({
          name: d.toLocaleDateString("id-ID", { weekday: "short" }),
          hadir: dayHadir,
          alfa: dAlfaTotal,
        });
      }
    }

    return {
      siswaCount,
      guruCount,
      hadirHariIni: hadir,
      terlambatHariIni: lambat,
      sakitCount: sakit,
      izinCount: izin,
      alfaCount: alfaTotal,
      chartData,
      siswaL,
      siswaP,
    };
  }, [
    students,
    teachers,
    attendance,
    settings,
    holidays,
    rekapDailyList,
    globalSiswaCount,
    globalGuruCount,
    globalSiswaL,
    globalSiswaP,
  ]);

  // Reporting states
  const [rekapFilter, setRekapFilter] = useState({
    bulan: new Date().toISOString().slice(0, 7),
    kelas: "",
    type: "Siswa",
  });

  useEffect(() => {
    if (
      session?.kelas &&
      (session?.role === "Guru Wali Kelas" || session?.isWali)
    ) {
      setRekapFilter((prev) => ({ ...prev, kelas: session.kelas }));
      setFilterClassAbsensi(session.kelas);
    }
  }, [session]);

  useEffect(() => {
    setPagination((prev) => ({ ...prev, rekapSiswa: 0, rekapGuru: 0 }));
  }, [rekapFilter.bulan, rekapFilter.kelas, rekapFilter.type]);

  const [editingProfileData, setEditingProfileData] = useState<any>(null);
  const [isUpdatingProfile, setIsUpdatingProfile] = useState(false);

  // Optimized Maps for Rekap
  const monthAttendanceMap = useMemo(() => {
    const map: Record<string, Attendance[]> = {};
    attendance.forEach((a) => {
      if (a.tanggal.startsWith(rekapFilter.bulan)) {
        if (!map[a.nisn]) map[a.nisn] = [];
        map[a.nisn].push(a);
      }
    });
    return map;
  }, [attendance, rekapFilter.bulan]);

  const teacherAttendanceMap = useMemo(() => {
    const map: Record<string, TeacherAttendance[]> = {};
    teacherAttendance.forEach((ta) => {
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
      refreshMasterData();
      alert("Profil berhasil diperbarui.");
      // Refresh teacher state if needed
    } catch (err) {
      alert("Gagal memperbarui profil.");
    } finally {
      setIsUpdatingProfile(false);
    }
  };

  const [searchTeacherReport, setSearchTeacherReport] = useState("");

  const handleOpenRoster = async () => {
    // Open a blank new window synchronously first to bypass browser popup blockers
    const newWindow = window.open("about:blank", "_blank");
    if (newWindow) {
      newWindow.document.write(`
        <html>
          <head>
            <title>Menghubungkan ke Roster...</title>
            <style>
              body {
                font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
                display: flex;
                flex-direction: column;
                align-items: center;
                justify-content: center;
                height: 100vh;
                margin: 0;
                background-color: #0c0a09;
                color: #f5f5f4;
              }
              .spinner {
                border: 4px solid rgba(255, 255, 255, 0.1);
                border-left-color: #22c55e;
                border-radius: 50%;
                width: 40px;
                height: 40px;
                animation: spin 1s linear infinite;
                margin-bottom: 20px;
              }
              @keyframes spin {
                to { transform: rotate(360deg); }
              }
              h1 {
                font-size: 18px;
                font-weight: 700;
                margin: 0 0 8px 0;
              }
              p {
                font-size: 13px;
                color: #a8a29e;
                margin: 0;
              }
            </style>
          </head>
          <body>
            <div class="spinner"></div>
            <h1>Menghubungkan ke SIGAP Roster</h1>
            <p>Mohon tunggu sebentar, sedang menyinkronkan token keamanan...</p>
          </body>
        </html>
      `);
    }

    toggleLoader(true);
    try {
      const token = await firestoreService.generateOneTimeToken(session);
      const url = `https://criet-roster.vercel.app/?token=${token}&role=${session?.role || "Guru"}&name=${encodeURIComponent(session?.name || "User")}&ott=${token}&ts=${Date.now()}`;
      if (newWindow) {
        newWindow.location.href = url;
      } else {
        const opened = window.open(url, "_blank", "noopener,noreferrer");
        if (!opened) {
          alert(
            "Pop-up diblokir! Silakan izinkan pop-up di browser Anda agar dapat membuka aplikasi Roster.",
          );
        }
      }
    } catch (err) {
      console.error(err);
      if (newWindow) {
        newWindow.close();
      }
      alert("Gagal membuat token login sekali pakai.");
    } finally {
      toggleLoader(false);
    }
  };

  const renderRoster = () => {
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
            Hasilkan jadwal mengajar (roster) sekolah secara otomatis dan
            efisien menggunakan teknologi AI. Klik tombol di bawah untuk
            diarahkan ke platform pembuatan roster.
          </p>

          <div className="space-y-4">
            <motion.button
              whileHover={{ scale: 1.02 }}
              whileTap={{ scale: 0.98 }}
              onClick={handleOpenRoster}
              className="w-full bg-zinc-900 text-white rounded-[2rem] py-6 px-8 font-black uppercase tracking-widest shadow-xl hover:bg-zinc-800 transition-all flex items-center justify-center gap-4 group cursor-pointer border-0"
            >
              <ExternalLink
                size={24}
                className="text-green-400 group-hover:rotate-12 transition-transform"
              />
              Buka Dashboard Roster
            </motion.button>

            <div className="flex items-center justify-center gap-2 py-3 px-6 bg-zinc-50 rounded-2xl border border-zinc-100">
              <ShieldCheck size={16} className="text-green-600" />
              <span className="text-[10px] font-black text-zinc-400 uppercase tracking-widest">
                Akses Terproteksi (Sesi Sekali Pakai)
              </span>
            </div>
          </div>

          <p className="mt-8 text-[11px] font-bold text-zinc-400 italic">
            *Anda akan diarahkan ke tab baru secara aman dengan Token Sekali
            Pakai. Link ini tidak dapat disalin atau dibuka ulang di perangkat
            lain.
          </p>
        </div>
      </div>
    );
  };

  const handleOpenKBC = async () => {
    // Open a blank new window synchronously first to bypass browser popup blockers
    const newWindow = window.open("about:blank", "_blank");
    if (newWindow) {
      newWindow.document.write(`
        <html>
          <head>
            <title>Menghubungkan ke KBC...</title>
            <style>
              body {
                font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
                display: flex;
                flex-direction: column;
                align-items: center;
                justify-content: center;
                height: 100vh;
                margin: 0;
                background-color: #0c0a09;
                color: #f5f5f4;
              }
              .spinner {
                border: 4px solid rgba(255, 255, 255, 0.1);
                border-left-color: #e11d48;
                border-radius: 50%;
                width: 40px;
                height: 40px;
                animation: spin 1s linear infinite;
                margin-bottom: 20px;
              }
              @keyframes spin {
                to { transform: rotate(360deg); }
              }
              h1 {
                font-size: 18px;
                font-weight: 700;
                margin: 0 0 8px 0;
              }
              p {
                font-size: 13px;
                color: #a8a29e;
                margin: 0;
              }
            </style>
          </head>
          <body>
            <div class="spinner"></div>
            <h1>Menghubungkan ke SIGAP KBC</h1>
            <p>Mohon tunggu sebentar, sedang menyinkronkan token keamanan...</p>
          </body>
        </html>
      `);
    }

    toggleLoader(true);
    try {
      const token = await firestoreService.generateOneTimeToken(session);
      const url = `https://kbc-mtsn2bombana.vercel.app/?token=${token}&role=${session?.role || "Guru"}&name=${encodeURIComponent(session?.name || "User")}&ott=${token}&ts=${Date.now()}`;
      if (newWindow) {
        newWindow.location.href = url;
      } else {
        const opened = window.open(url, "_blank", "noopener,noreferrer");
        if (!opened) {
          alert(
            "Pop-up diblokir! Silakan izinkan pop-up di browser Anda agar dapat membuka aplikasi KBC.",
          );
        }
      }
    } catch (err) {
      console.error(err);
      if (newWindow) {
        newWindow.close();
      }
      alert("Gagal membuat token login sekali pakai.");
    } finally {
      toggleLoader(false);
    }
  };

  const renderKBC = () => {
    const isAuthorized = session?.role === "Guru" || session?.role === "Admin";

    if (!isAuthorized) {
      return (
        <div className="flex flex-col items-center justify-center min-h-[500px] bg-white rounded-[3rem] p-12 shadow-2xl border border-red-50 text-center relative overflow-hidden">
          <div className="absolute top-0 left-0 w-full h-2 bg-gradient-to-r from-red-400 via-pink-500 to-rose-600"></div>
          <div className="absolute -top-24 -right-24 w-64 h-64 bg-red-50 rounded-full blur-3xl opacity-50"></div>
          <div className="absolute -bottom-24 -left-24 w-64 h-64 bg-rose-50 rounded-full blur-3xl opacity-50"></div>

          <div className="relative z-10 max-w-lg mx-auto">
            <motion.div
              initial={{ scale: 0.8, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              className="w-24 h-24 bg-gradient-to-br from-red-500 to-rose-600 rounded-[2.5rem] flex items-center justify-center mb-8 mx-auto shadow-lg shadow-red-200"
            >
              <AlertCircle className="text-white" size={44} />
            </motion.div>

            <h3 className="text-3xl font-black text-gray-900 uppercase tracking-tighter mb-4">
              Akses <span className="text-red-600">Ditolak</span>
            </h3>

            <p className="text-sm font-bold text-gray-500 mb-10 leading-relaxed">
              Anda tidak memiliki izin/hak akses untuk membuka halaman SIGAP
              KBC. Sebutkan akun guru atau admin aktif yang sah untuk
              melanjutkan.
            </p>
          </div>
        </div>
      );
    }

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
            <motion.button
              whileHover={{ scale: 1.02 }}
              whileTap={{ scale: 0.98 }}
              onClick={handleOpenKBC}
              className="w-full bg-zinc-900 text-white rounded-[2rem] py-6 px-8 font-black uppercase tracking-widest shadow-xl hover:bg-zinc-800 transition-all flex items-center justify-center gap-4 group cursor-pointer border-0"
            >
              <ExternalLink
                size={24}
                className="text-rose-400 group-hover:rotate-12 transition-transform"
              />
              Buka Aplikasi KBC
            </motion.button>

            <div className="flex items-center justify-center gap-2 py-3 px-6 bg-zinc-50 rounded-2xl border border-zinc-100">
              <ShieldCheck size={16} className="text-green-600" />
              <span className="text-[10px] font-black text-zinc-400 uppercase tracking-widest">
                Akses Terproteksi (Sesi Sekali Pakai)
              </span>
            </div>
          </div>

          <p className="mt-8 text-[11px] font-bold text-zinc-400 italic">
            *Anda akan diarahkan ke tab baru secara aman dengan Token Sekali
            Pakai. Link ini tidak dapat disalin atau dibuka ulang di perangkat
            lain.
          </p>
        </div>
      </div>
    );
  };

  const renderLaporanCapaianGuru = () => {
    const selectedTeacher = teachers.find(
      (t) => t.nip === selectedTeacherNipForCapaian,
    );

    return (
      <div className="space-y-6">
        {!selectedTeacher ? (
          <div className="bg-white rounded-[2rem] p-8 shadow-xl border border-gray-100">
            <div className="flex flex-col md:flex-row justify-between md:items-center gap-4 mb-8">
              <h3 className="text-xl font-black text-green-950 uppercase tracking-widest flex items-center gap-3">
                <User className="text-green-600" /> Pilih Guru
              </h3>
              <div className="relative w-full md:w-64">
                <Search
                  className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-300"
                  size={16}
                />
                <input
                  type="text"
                  placeholder="Cari nama guru..."
                  value={searchTeacherReport}
                  onChange={(e) => setSearchTeacherReport(e.target.value)}
                  className="w-full pl-10 pr-4 py-2 bg-gray-50 border-0 rounded-xl text-sm font-bold focus:ring-2 focus:ring-green-500"
                />
              </div>
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
              {teachers
                .filter((t) => t.nip !== "ADMIN001" && t.role !== "Admin")
                .filter((t) =>
                  t.nama
                    .toLowerCase()
                    .includes(searchTeacherReport.toLowerCase()),
                )
                .map((t) => (
                  <button
                    key={t.nip}
                    onClick={() => setSelectedTeacherNipForCapaian(t.nip)}
                    className="bg-gray-50 hover:bg-green-50 p-6 rounded-[2rem] border border-gray-100 text-left transition-all hover:shadow-md flex items-center gap-4 group"
                  >
                    <div className="w-12 h-12 rounded-2xl bg-zinc-900 flex items-center justify-center text-white overflow-hidden shadow-sm">
                      {t.foto ? (
                        <img
                          src={t.foto}
                          className="w-full h-full object-cover"
                          alt={t.nama}
                        />
                      ) : (
                        <User size={20} />
                      )}
                    </div>
                    <div className="flex-1">
                      <p className="text-sm font-black text-zinc-900 group-hover:text-green-900 transition-colors">
                        {t.nama}
                      </p>
                      <p className="text-[10px] font-black text-zinc-400 uppercase tracking-widest">
                        {t.jabatan || "Guru"}
                      </p>
                    </div>
                    <ChevronRight
                      size={16}
                      className="text-zinc-300 group-hover:text-green-600"
                    />
                  </button>
                ))}
            </div>
          </div>
        ) : (
          <motion.div
            initial={{ opacity: 0, x: 20 }}
            animate={{ opacity: 1, x: 0 }}
            className="space-y-6 pb-20"
          >
            <div className="flex items-center gap-4">
              <button
                onClick={() => setSelectedTeacherNipForCapaian(null)}
                className="bg-white p-3 rounded-2xl shadow-sm border border-gray-100 text-zinc-400 hover:text-green-900 transition-colors"
              >
                <ArrowLeft size={20} />
              </button>
              <div>
                <h3 className="font-black text-xl text-green-950">
                  {selectedTeacher.nama}
                </h3>
                <p className="text-xs font-bold text-gray-400 uppercase tracking-widest">
                  Laporan Capaian Mengajar
                </p>
              </div>
            </div>

            <div className="bg-white rounded-[2.5rem] p-8 md:p-12 shadow-xl border border-gray-100">
              <h3 className="text-lg font-black text-green-950 mb-6 flex items-center gap-3 uppercase tracking-widest">
                <Calendar className="text-green-600" /> Jam Mengajar & Laporan
                Capaian
              </h3>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                {(() => {
                  const now = new Date();
                  const year = now.getFullYear();
                  const month = now.getMonth();
                  const currentMonthPrefix = now.toISOString().slice(0, 7);

                  const mySchedules = teachingSchedules.filter(
                    (ts) => ts.nip === selectedTeacher.nip,
                  );

                  // Group schedules by Class and Mapel
                  const groupedSchedules: {
                    [key: string]: {
                      kelas: string;
                      mapel: string;
                      days: string[];
                      sessions: any[];
                    };
                  } = {};

                  mySchedules.forEach((s) => {
                    const key = `${s.kelas}-${s.mapel || "unnamed"}`;
                    if (!groupedSchedules[key]) {
                      groupedSchedules[key] = {
                        kelas: s.kelas,
                        mapel: s.mapel || "-",
                        days: [],
                        sessions: [],
                      };
                    }
                    if (!groupedSchedules[key].days.includes(s.hari)) {
                      groupedSchedules[key].days.push(s.hari);
                    }
                    groupedSchedules[key].sessions.push(s);
                  });

                  const sortedGroups = Object.values(groupedSchedules);

                  return sortedGroups.length > 0 ? (
                    sortedGroups.map((group, idx) => {
                      let totalMonthlyTarget = 0;
                      let totalActualSessions = 0;

                      group.sessions.forEach((s) => {
                        totalMonthlyTarget +=
                          getEffectiveTargetSessionsForSchedule(
                            s,
                            year,
                            month,
                            holidays,
                            settings,
                          );

                        const actual = teacherAttendance.filter(
                          (ta) =>
                            ta.nip === selectedTeacher.nip &&
                            ta.kelas === s.kelas &&
                            ta.tanggal.startsWith(currentMonthPrefix) &&
                            (() => {
                              const [y, m, d] = ta.tanggal
                                .split("-")
                                .map(Number);
                              const dObj = new Date(y, m - 1, d);
                              const dayNames = [
                                "Minggu",
                                "Senin",
                                "Selasa",
                                "Rabu",
                                "Kamis",
                                "Jumat",
                                "Sabtu",
                              ];
                              return dayNames[dObj.getDay()] === s.hari;
                            })(),
                        ).length;
                        totalActualSessions += actual;
                      });

                      const achievementPerc =
                        totalMonthlyTarget > 0
                          ? Math.min(
                              100,
                              (totalActualSessions / totalMonthlyTarget) * 100,
                            )
                          : 0;

                      return (
                        <div
                          key={idx}
                          className="bg-white p-6 rounded-[2rem] border border-gray-100 shadow-sm hover:shadow-md transition-shadow relative overflow-hidden group"
                        >
                          <div className="absolute top-0 right-0 p-4">
                            <div
                              className={`w-10 h-10 rounded-2xl flex items-center justify-center ${achievementPerc >= 100 ? "bg-green-100 text-green-600" : "bg-orange-100 text-orange-600"}`}
                            >
                              <Award size={20} />
                            </div>
                          </div>

                          <div className="mb-4">
                            <span className="text-[10px] font-black text-green-600 bg-green-50 px-3 py-1 rounded-full uppercase tracking-widest">
                              {group.days.join(" / ")}
                            </span>
                          </div>

                          <div className="space-y-1 mb-6">
                            <h4 className="text-xl font-black text-zinc-900">
                              {group.kelas}
                            </h4>
                            <p className="text-sm font-bold text-gray-400">
                              {group.mapel}
                            </p>
                          </div>

                          <div className="grid grid-cols-2 gap-4 pt-4 border-t border-gray-50">
                            <div>
                              <p className="text-[9px] font-black text-gray-400 uppercase tracking-widest mb-1">
                                Target
                              </p>
                              <div className="flex items-baseline gap-1">
                                <span className="text-lg font-black text-zinc-900">
                                  {totalMonthlyTarget}
                                </span>
                                <span className="text-[10px] font-bold text-gray-400">
                                  Sesi
                                </span>
                              </div>
                            </div>
                            <div>
                              <p className="text-[9px] font-black text-gray-400 uppercase tracking-widest mb-1">
                                Capaian
                              </p>
                              <div className="flex items-baseline gap-1">
                                <span className="text-lg font-black text-green-700">
                                  {totalActualSessions}
                                </span>
                                <span className="text-[10px] font-bold text-gray-400">
                                  Sesi
                                </span>
                              </div>
                            </div>
                          </div>

                          <div className="mt-4">
                            <div className="flex justify-between items-center mb-1.5">
                              <span className="text-[9px] font-black text-gray-400 uppercase tracking-widest">
                                Progress Bulanan
                              </span>
                              <span className="text-[10px] font-black text-green-700">
                                {Math.round(achievementPerc)}%
                              </span>
                            </div>
                            <div className="h-2 w-full bg-gray-100 rounded-full overflow-hidden text-[0px]">
                              <motion.div
                                initial={{ width: 0 }}
                                animate={{ width: `${achievementPerc}%` }}
                                transition={{ duration: 1, ease: "easeOut" }}
                                className={`h-full rounded-full ${achievementPerc >= 100 ? "bg-green-500 shadow-[0_0_8px_rgba(34,197,94,0.4)]" : "bg-orange-500"}`}
                              />
                            </div>
                          </div>
                        </div>
                      );
                    })
                  ) : (
                    <div className="col-span-full p-12 text-center bg-gray-50 rounded-[2rem] border border-dashed border-gray-200">
                      <p className="text-sm font-bold text-gray-400">
                        Guru ini belum memiliki target jam mengajar.
                      </p>
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

  const getIndonesianDayName = (dateStr: string) => {
    const d = new Date(dateStr);
    const days = [
      "Minggu",
      "Senin",
      "Selasa",
      "Rabu",
      "Kamis",
      "Jumat",
      "Sabtu",
    ];
    return days[d.getDay()];
  };

  const getDatesInRange = (startDate: string, endDate: string) => {
    const dates = [];
    let curr = new Date(startDate);
    const end = new Date(endDate);
    while (curr <= end) {
      dates.push(curr.toISOString().split("T")[0]);
      curr.setDate(curr.getDate() + 1);
    }
    return dates;
  };

  const submitManualTeacherAttendance = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!manT_nip) {
      alert("Pilih guru terlebih dahulu.");
      return;
    }
    const targetTeacher = teachers.find((t) => t.nip === manT_nip);
    if (!targetTeacher) {
      alert("Guru tidak ditemukan.");
      return;
    }

    let datesToRecord: string[] = [];
    if (manT_mode === "single") {
      if (!manT_date) {
        alert("Pilih tanggal terlebih dahulu.");
        return;
      }
      datesToRecord = [manT_date];
    } else {
      if (!manT_startDate || !manT_endDate) {
        alert("Pilih rentang tanggal lengkap.");
        return;
      }
      if (manT_startDate > manT_endDate) {
        alert("Tanggal mulai tidak boleh melebihi tanggal selesai.");
        return;
      }
      datesToRecord = getDatesInRange(manT_startDate, manT_endDate);
    }

    if (!manT_keterangan.trim()) {
      alert("Isi keterangan izin/sakit terlebih dahulu.");
      return;
    }

    setManT_saving(true);
    toggleLoader(true);

    try {
      let savedCount = 0;
      for (const tanggal of datesToRecord) {
        const indonesianDay = getIndonesianDayName(tanggal);
        const matchingSchedules = teachingSchedules.filter(
          (s) => s.nip === manT_nip && s.hari === indonesianDay,
        );

        if (matchingSchedules.length > 0) {
          for (const s of matchingSchedules) {
            const record = {
              id: `${manT_nip}-${s.kelas}-${tanggal}`,
              nip: manT_nip,
              nama: targetTeacher.nama,
              kelas: s.kelas,
              mapel: s.mapel || "-",
              tanggal,
              jam: "-",
              terlambat: 0,
              status: manT_status,
              keterangan: manT_keterangan,
            };
            await firestoreService.saveTeacherAttendanceManual(record);
            savedCount++;
          }
        } else {
          const record = {
            id: `${manT_nip}-general-${tanggal}`,
            nip: manT_nip,
            nama: targetTeacher.nama,
            kelas: "-",
            mapel: "-",
            tanggal,
            jam: "-",
            terlambat: 0,
            status: manT_status,
            keterangan: manT_keterangan,
          };
          await firestoreService.saveTeacherAttendanceManual(record);
          savedCount++;
        }
      }

      triggerSuccess(
        "BERHASIL",
        `Sukses menyimpan ${savedCount} rekaman presensi manual untuk ${targetTeacher.nama}.`,
      );
      setManT_nip("");
      setManT_date("");
      setManT_startDate("");
      setManT_endDate("");
      setManT_keterangan("");
    } catch (err) {
      console.error(err);
      alert("Terjadi kesalahan saat menyimpan.");
    } finally {
      setManT_saving(false);
      toggleLoader(false);
    }
  };

  const renderInputKehadiranGuru = () => {
    const manualEntries = teacherAttendance
      .filter(
        (ta) =>
          ta.status && ["Sakit", "Izin", "Alfa", "Alpa"].includes(ta.status),
      )
      .sort((a, b) => b.tanggal.localeCompare(a.tanggal));

    const totalKeteranganCount = manualEntries.length;
    const activeTeachers = teachers.filter(
      (t) =>
        t.jabatan !== "Kamad" && t.nip !== "ADMIN001" && t.role !== "Admin",
    );

    return (
      <div className="space-y-6 max-w-7xl mx-auto px-4 md:px-0">
        <div className="flex flex-col">
          <h2 className="text-xl font-bold flex items-center gap-2 text-zinc-900">
            <CalendarRange size={24} className="text-green-700" /> Input
            Kehadiran Guru
          </h2>
          <p className="text-xs text-zinc-400 font-semibold">
            Form khusus Kamad untuk mencatat izin, sakit, dan ketidakhadiran
            guru secara manual
          </p>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 items-start">
          <div className="lg:col-span-12 xl:col-span-5 bg-white p-6 rounded-3xl border border-gray-100 shadow-sm space-y-6">
            <h3 className="text-sm font-black text-gray-800 uppercase tracking-wider mb-4 border-b border-gray-50 pb-2">
              Form Presensi Manual
            </h3>
            <form
              onSubmit={submitManualTeacherAttendance}
              className="space-y-4"
            >
              <div>
                <label className="block text-[10px] font-black text-gray-400 uppercase tracking-widest mb-1">
                  Pilih Guru
                </label>
                <select
                  value={manT_nip}
                  onChange={(e) => setManT_nip(e.target.value)}
                  className="w-full bg-gray-50 border-0 rounded-xl px-4 py-3 text-sm font-semibold text-zinc-800 focus:ring-2 focus:ring-green-600 transition-all cursor-pointer"
                  required
                >
                  <option value="">-- Pilih Guru --</option>
                  {activeTeachers.map((t) => (
                    <option key={t.nip} value={t.nip}>
                      {t.nama}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-[10px] font-black text-gray-400 uppercase tracking-widest mb-2">
                  Pilih Status Kehadiran
                </label>
                <div className="grid grid-cols-2 gap-2">
                  <button
                    type="button"
                    onClick={() => setManT_status("Sakit")}
                    className={`p-3 rounded-xl border text-center transition-all cursor-pointer ${
                      manT_status === "Sakit"
                        ? "border-orange-500 bg-orange-50/50 text-orange-700 font-extrabold ring-2 ring-orange-100"
                        : "border-gray-100 bg-gray-50 hover:bg-gray-100 text-gray-500 text-xs font-semibold"
                    }`}
                  >
                    Sakit
                  </button>
                  <button
                    type="button"
                    onClick={() => setManT_status("Izin")}
                    className={`p-3 rounded-xl border text-center transition-all cursor-pointer ${
                      manT_status === "Izin"
                        ? "border-amber-500 bg-amber-50/50 text-amber-700 font-extrabold ring-2 ring-amber-100"
                        : "border-gray-100 bg-gray-50 hover:bg-gray-100 text-gray-500 text-xs font-semibold"
                    }`}
                  >
                    Izin
                  </button>
                </div>
              </div>

              <div>
                <label className="block text-[10px] font-black text-gray-400 uppercase tracking-widest mb-1">
                  Metode Penginputan Tanggal
                </label>
                <div className="flex gap-2 p-1 bg-gray-100 rounded-xl w-full">
                  <button
                    type="button"
                    onClick={() => setManT_mode("single")}
                    className={`flex-1 py-1.5 text-xs font-black rounded-lg transition-all uppercase tracking-wider ${
                      manT_mode === "single"
                        ? "bg-green-800 text-white shadow"
                        : "text-gray-400 font-bold"
                    }`}
                  >
                    Satu Hari
                  </button>
                  <button
                    type="button"
                    onClick={() => setManT_mode("range")}
                    className={`flex-1 py-1.5 text-xs font-black rounded-lg transition-all uppercase tracking-wider ${
                      manT_mode === "range"
                        ? "bg-green-800 text-white shadow"
                        : "text-gray-400 font-bold"
                    }`}
                  >
                    Rentang Tanggal
                  </button>
                </div>
              </div>

              {manT_mode === "single" ? (
                <div>
                  <label className="block text-[10px] font-black text-gray-400 uppercase tracking-widest mb-1">
                    Tanggal
                  </label>
                  <input
                    type="date"
                    value={manT_date}
                    onChange={(e) => setManT_date(e.target.value)}
                    className="w-full bg-gray-50 border-0 rounded-xl px-4 py-3 text-sm font-semibold text-zinc-800 focus:ring-2 focus:ring-green-600 transition-all cursor-pointer"
                    required={manT_mode === "single"}
                  />
                </div>
              ) : (
                <div className="grid grid-cols-2 gap-2">
                  <div>
                    <label className="block text-[10px] font-black text-gray-400 uppercase tracking-widest mb-1">
                      Mulai
                    </label>
                    <input
                      type="date"
                      value={manT_startDate}
                      onChange={(e) => setManT_startDate(e.target.value)}
                      className="w-full bg-gray-50 border-0 rounded-xl px-3 py-3 text-xs font-semibold text-zinc-800 focus:ring-2 focus:ring-green-600 transition-all cursor-pointer"
                      required={manT_mode === "range"}
                    />
                  </div>
                  <div>
                    <label className="block text-[10px] font-black text-gray-400 uppercase tracking-widest mb-1">
                      Selesai
                    </label>
                    <input
                      type="date"
                      value={manT_endDate}
                      onChange={(e) => setManT_endDate(e.target.value)}
                      className="w-full bg-gray-50 border-0 rounded-xl px-3 py-3 text-xs font-semibold text-zinc-800 focus:ring-2 focus:ring-green-600 transition-all cursor-pointer"
                      required={manT_mode === "range"}
                    />
                  </div>
                </div>
              )}

              <div>
                <label className="block text-[10px] font-black text-gray-400 uppercase tracking-widest mb-1">
                  Keterangan / Alasan
                </label>
                <textarea
                  value={manT_keterangan}
                  onChange={(e) => setManT_keterangan(e.target.value)}
                  placeholder={`Masukkan alasan guru ${manT_status.toLowerCase()} (misal: sakit flu, izin dinas luar, dll)`}
                  rows={3}
                  className="w-full bg-gray-50 border-0 rounded-xl px-4 py-3 text-sm font-semibold text-zinc-800 focus:ring-2 focus:ring-green-600 transition-all"
                  required
                />
              </div>

              <button
                type="submit"
                disabled={manT_saving}
                className="w-full bg-green-800 hover:bg-green-700 text-white font-black uppercase tracking-wider text-xs py-4 rounded-xl shadow-md transition-all flex items-center justify-center gap-2 cursor-pointer disabled:opacity-50"
              >
                {manT_saving ? "Menyimpan..." : "Simpan Kehadiran"}
              </button>
            </form>
          </div>

          <div className="lg:col-span-12 xl:col-span-7 bg-white p-6 rounded-3xl border border-gray-100 shadow-sm space-y-4">
            <div className="flex justify-between items-center border-b border-gray-50 pb-2">
              <h3 className="text-sm font-black text-gray-800 uppercase tracking-wider">
                Histori Presensi Manual
              </h3>
              <span className="bg-green-100 text-green-800 px-2.5 py-0.5 rounded-full text-[10px] font-black uppercase tracking-wider">
                {totalKeteranganCount} Record
              </span>
            </div>

            <div className="relative">
              <span className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none text-gray-400">
                <Search size={14} />
              </span>
              <input
                type="text"
                placeholder="Cari nama guru..."
                value={manT_searchTerm}
                onChange={(e) => setManT_searchTerm(e.target.value)}
                className="w-full pl-10 pr-4 py-2.5 bg-gray-50 border-0 rounded-xl text-xs font-semibold text-zinc-800 focus:ring-2 focus:ring-green-600 transition-all"
              />
            </div>

            <div className="overflow-x-auto rounded-xl border border-gray-100 max-h-[450px] overflow-y-auto">
              <table className="w-full text-left text-xs min-w-[500px]">
                <thead className="bg-gray-50 text-gray-400 uppercase font-black text-[10px] tracking-wider sticky top-0 z-10 border-b border-gray-100">
                  <tr>
                    <th className="px-4 py-3 text-center w-12">No</th>
                    <th className="px-4 py-3">Nama Guru</th>
                    <th className="px-4 py-3 text-center">Tanggal</th>
                    <th className="px-4 py-3 text-center">Kelas</th>
                    <th className="px-4 py-3 text-center">Status</th>
                    <th className="px-4 py-3">Keterangan</th>
                    <th className="px-4 py-3 text-center w-16">Aksi</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-50">
                  {manualEntries
                    .filter((entry) =>
                      entry.nama
                        .toLowerCase()
                        .includes(manT_searchTerm.toLowerCase()),
                    )
                    .map((item, idx) => {
                      let badgeColor = "";
                      if (item.status === "Sakit")
                        badgeColor =
                          "bg-orange-100 text-orange-700 border-orange-200";
                      else if (item.status === "Izin")
                        badgeColor =
                          "bg-amber-100 text-amber-700 border-amber-200";
                      else
                        badgeColor = "bg-red-100 text-red-700 border-red-200";

                      return (
                        <tr
                          key={item.id}
                          className="hover:bg-gray-50 transition-colors"
                        >
                          <td className="px-4 py-3 text-center font-bold text-gray-400">
                            {idx + 1}
                          </td>
                          <td className="px-4 py-3 font-bold text-zinc-800 whitespace-nowrap">
                            {item.nama}
                          </td>
                          <td className="px-4 py-3 text-center font-medium text-gray-500 whitespace-nowrap">
                            {formatIndoDate(item.tanggal)}
                          </td>
                          <td className="px-4 py-3 text-center font-bold">
                            <span className="bg-gray-100 text-gray-600 px-1.5 py-0.5 rounded text-[9px]">
                              {item.kelas}
                            </span>
                          </td>
                          <td className="px-4 py-3 text-center font-black">
                            <span
                              className={`px-2 py-0.5 rounded-full border text-[9px] ${badgeColor}`}
                            >
                              {item.status}
                            </span>
                          </td>
                          <td
                            className="px-4 py-3 text-zinc-600 max-w-[150px] truncate"
                            title={item.keterangan || "-"}
                          >
                            {item.keterangan || "-"}
                          </td>
                          <td className="px-4 py-3 text-center">
                            <button
                              onClick={() => {
                                setConfirmModal({
                                  show: true,
                                  title: "Hapus Kehadiran Manual?",
                                  entityName: `Izin/Sakit: ${item.nama} pada ${item.tanggal}`,
                                  message:
                                    "Data ketidakhadiran guru ini akan dihapus dari sistem.",
                                  onConfirm: async () => {
                                    toggleLoader(true);
                                    try {
                                      await firestoreService.hapusAbsensiGuru(
                                        item.id,
                                      );
                                      triggerSuccess(
                                        "BERHASIL",
                                        "Presensi manual berhasil dihapus.",
                                      );
                                    } catch (err) {
                                      alert("Gagal menghapus data.");
                                    } finally {
                                      toggleLoader(false);
                                    }
                                  },
                                });
                              }}
                              className="text-red-500 hover:bg-red-50 p-1.5 rounded-lg transition-all cursor-pointer"
                            >
                              <Trash2 size={14} />
                            </button>
                          </td>
                        </tr>
                      );
                    })}
                  {manualEntries.filter((entry) =>
                    entry.nama
                      .toLowerCase()
                      .includes(manT_searchTerm.toLowerCase()),
                  ).length === 0 && (
                    <tr>
                      <td
                        colSpan={7}
                        className="text-center py-8 text-gray-400 font-bold italic"
                      >
                        Belum ada data izin/sakit guru.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      </div>
    );
  };

  const renderProfile = () => {
    const teacher = teachers.find((t) => t.nip === session?.uid);
    const data = editingProfileData || teacher;

    if (!data)
      return (
        <div className="p-8 text-center text-gray-400">
          Memuat data profil...
        </div>
      );

    const mySchedules = teachingSchedules.filter((ts) => ts.nip === data.nip);

    return (
      <motion.div
        initial={{ opacity: 0, y: 20 }}
        animate={{ opacity: 1, y: 0 }}
        className="max-w-4xl mx-auto space-y-8 pb-20"
      >
        <div className="bg-white rounded-[2.5rem] p-8 md:p-12 shadow-xl border border-gray-100 overflow-hidden relative">
          <div className="absolute top-0 left-0 w-full h-32 bg-green-900/5 -z-10" />

          <div className="flex flex-col items-center mb-10">
            <div className="relative mb-6">
              <div className="w-32 h-32 rounded-[2.5rem] bg-green-900 flex items-center justify-center text-white overflow-hidden shadow-2xl ring-4 ring-white">
                {data.foto ? (
                  <img
                    src={data.foto}
                    className="w-full h-full object-cover"
                    crossOrigin="anonymous"
                    alt="Foto Profil"
                  />
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
                        setEditingProfileData({
                          ...data,
                          foto: reader.result as string,
                        });
                      };
                      reader.readAsDataURL(file);
                    }
                  }}
                />
              </label>
            </div>
            <h2 className="text-2xl font-black text-green-950 mb-1">
              {data.nama}
            </h2>
            <p className="text-xs font-bold text-green-600 uppercase tracking-widest bg-green-50 px-4 py-1.5 rounded-full">
              {data.jabatan || "Guru"}
            </p>
          </div>

          <form onSubmit={handleUpdateProfile} className="space-y-6">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
              <div className="space-y-2">
                <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest ml-1">
                  NIP / ID
                </label>
                <input
                  type="text"
                  readOnly
                  value={data.nip}
                  className="w-full bg-gray-50 border-0 rounded-2xl px-6 py-4 text-sm font-bold text-gray-400 cursor-not-allowed"
                />
              </div>
              <div className="space-y-2">
                <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest ml-1">
                  Nama Lengkap
                </label>
                <input
                  type="text"
                  required
                  value={data.nama}
                  onChange={(e) =>
                    setEditingProfileData({ ...data, nama: e.target.value })
                  }
                  className="w-full bg-gray-50 focus:bg-white focus:ring-2 focus:ring-green-100 border-0 rounded-2xl px-6 py-4 text-sm font-bold transition-all"
                />
              </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-6 p-6 bg-zinc-50 rounded-[2rem] border border-zinc-100">
              <div className="space-y-2">
                <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">
                  Username Login
                </label>
                <input
                  type="text"
                  required
                  value={data.user}
                  onChange={(e) =>
                    setEditingProfileData({ ...data, user: e.target.value })
                  }
                  className="w-full bg-white border border-zinc-200 rounded-2xl px-6 py-4 text-sm font-bold shadow-sm"
                />
              </div>
              <div className="space-y-2">
                <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">
                  Password Login
                </label>
                <div className="relative">
                  <input
                    type={showPass ? "text" : "password"}
                    required
                    value={data.pass}
                    onChange={(e) =>
                      setEditingProfileData({ ...data, pass: e.target.value })
                    }
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
            <Calendar className="text-green-600" /> Jam Mengajar & Laporan
            Capaian
          </h3>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            {(() => {
              const now = new Date();
              const year = now.getFullYear();
              const month = now.getMonth();
              const currentMonthPrefix = now.toISOString().slice(0, 7);

              // Group schedules by Class and Mapel
              const groupedSchedules: {
                [key: string]: {
                  kelas: string;
                  mapel: string;
                  days: string[];
                  sessions: any[];
                };
              } = {};

              mySchedules.forEach((s) => {
                const key = `${s.kelas}-${s.mapel || "unnamed"}`;
                if (!groupedSchedules[key]) {
                  groupedSchedules[key] = {
                    kelas: s.kelas,
                    mapel: s.mapel || "-",
                    days: [],
                    sessions: [],
                  };
                }
                if (!groupedSchedules[key].days.includes(s.hari)) {
                  groupedSchedules[key].days.push(s.hari);
                }
                groupedSchedules[key].sessions.push(s);
              });

              const sortedGroups = Object.values(groupedSchedules);

              return sortedGroups.length > 0 ? (
                sortedGroups.map((group, idx) => {
                  let totalMonthlyTarget = 0;
                  let totalActualSessions = 0;

                  group.sessions.forEach((s) => {
                    totalMonthlyTarget += getEffectiveTargetSessionsForSchedule(
                      s,
                      year,
                      month,
                      holidays,
                      settings,
                    );

                    const actual = teacherAttendance.filter(
                      (ta) =>
                        ta.nip === data.nip &&
                        ta.kelas === s.kelas &&
                        ta.tanggal.startsWith(currentMonthPrefix) &&
                        (() => {
                          const [y, m, d] = ta.tanggal.split("-").map(Number);
                          const dObj = new Date(y, m - 1, d);
                          const dayNames = [
                            "Minggu",
                            "Senin",
                            "Selasa",
                            "Rabu",
                            "Kamis",
                            "Jumat",
                            "Sabtu",
                          ];
                          return dayNames[dObj.getDay()] === s.hari;
                        })(),
                    ).length;
                    totalActualSessions += actual;
                  });

                  const achievementPerc =
                    totalMonthlyTarget > 0
                      ? Math.min(
                          100,
                          (totalActualSessions / totalMonthlyTarget) * 100,
                        )
                      : 0;

                  return (
                    <div
                      key={idx}
                      className="bg-white p-6 rounded-[2rem] border border-gray-100 shadow-sm hover:shadow-md transition-shadow relative overflow-hidden group"
                    >
                      <div className="absolute top-0 right-0 p-4">
                        <div
                          className={`w-10 h-10 rounded-2xl flex items-center justify-center ${achievementPerc >= 100 ? "bg-green-100 text-green-600" : "bg-orange-100 text-orange-600"}`}
                        >
                          <Award size={20} />
                        </div>
                      </div>

                      <div className="mb-4">
                        <span className="text-[10px] font-black text-green-600 bg-green-50 px-3 py-1 rounded-full uppercase tracking-widest">
                          {group.days.join(" / ")}
                        </span>
                      </div>

                      <div className="space-y-1 mb-6">
                        <h4 className="text-xl font-black text-zinc-900">
                          {group.kelas}
                        </h4>
                        <p className="text-sm font-bold text-gray-400">
                          {group.mapel}
                        </p>
                      </div>

                      <div className="grid grid-cols-2 gap-4 pt-4 border-t border-gray-50">
                        <div>
                          <p className="text-[9px] font-black text-gray-400 uppercase tracking-widest mb-1">
                            Target
                          </p>
                          <div className="flex items-baseline gap-1">
                            <span className="text-lg font-black text-zinc-900">
                              {totalMonthlyTarget}
                            </span>
                            <span className="text-[10px] font-bold text-gray-400">
                              Sesi
                            </span>
                          </div>
                        </div>
                        <div>
                          <p className="text-[9px] font-black text-gray-400 uppercase tracking-widest mb-1">
                            Capaian
                          </p>
                          <div className="flex items-baseline gap-1">
                            <span className="text-lg font-black text-green-700">
                              {totalActualSessions}
                            </span>
                            <span className="text-[10px] font-bold text-gray-400">
                              Sesi
                            </span>
                          </div>
                        </div>
                      </div>

                      <div className="mt-4">
                        <div className="flex justify-between items-center mb-1.5">
                          <span className="text-[9px] font-black text-gray-400 uppercase tracking-widest">
                            Progress Bulanan
                          </span>
                          <span className="text-[10px] font-black text-green-700">
                            {Math.round(achievementPerc)}%
                          </span>
                        </div>
                        <div className="h-2 w-full bg-gray-100 rounded-full overflow-hidden text-[0px]">
                          <motion.div
                            initial={{ width: 0 }}
                            animate={{ width: `${achievementPerc}%` }}
                            transition={{ duration: 1, ease: "easeOut" }}
                            className={`h-full rounded-full ${achievementPerc >= 100 ? "bg-green-500 shadow-[0_0_8px_rgba(34,197,94,0.4)]" : "bg-orange-500"}`}
                          />
                        </div>
                      </div>
                    </div>
                  );
                })
              ) : (
                <div className="col-span-full p-12 text-center bg-gray-50 rounded-[2rem] border border-dashed border-gray-200">
                  <p className="text-sm font-bold text-gray-400">
                    Belum ada target jam mengajar yang diatur oleh Admin.
                  </p>
                </div>
              );
            })()}
          </div>
          <p className="text-[10px] text-gray-400 mt-6 font-bold uppercase tracking-widest italic">
            *Data jam mengajar hanya dapat diubah oleh Administrator.
          </p>
        </div>

        <p className="text-center text-[10px] text-gray-400 font-bold uppercase tracking-widest">
          Pembaruan data profil anda akan sinkron ke dashboard admin secara
          otomatis.
        </p>
      </motion.div>
    );
  };

  // Login states
  const [loginRole, setLoginRole] = useState<"Siswa" | "Guru" | "Admin">(
    "Siswa",
  );
  const [loginNisn, setLoginNisn] = useState("");
  const [loginUser, setLoginUser] = useState("");
  const [loginPass, setLoginPass] = useState("");
  const [loginError, setLoginError] = useState("");
  const [showPass, setShowPass] = useState(false);

  // Scanner states
  const [scanType, setScanType] = useState<"Siswa" | "Kelas">("Siswa");
  const [scanResult, setScanResult] = useState<{
    success: boolean;
    message: string;
  } | null>(null);
  const [scanning, setScanning] = useState(false);
  const [classScanFailCount, setClassScanFailCount] = useState(0);

  useEffect(() => {
    let html5QrCode: any = null;

    const startScanner = async () => {
      if (activePanel === "scanner" && scanning) {
        // Wait a small bit for DOM to settle
        await new Promise((resolve) => setTimeout(resolve, 100));

        const element = document.getElementById("reader");
        if (!element) return;

        try {
          const { Html5Qrcode } = await import("html5-qrcode");
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
                if (navigator.onLine) {
                  try {
                    if (scanType === "Siswa") {
                      res = await firestoreService.processScan(decodedText);
                    } else {
                      res = await firestoreService.processClassScan(
                        session?.uid || "",
                        decodedText,
                      );
                    }
                  } catch (e) {
                    console.warn("Firestore scan error, falling back to offline mode:", e);
                    res = processOfflineScan(decodedText, scanType);
                  }
                } else {
                  res = processOfflineScan(decodedText, scanType);
                }

                if (res.success) {
                  const audio = new Audio(
                    "https://assets.mixkit.co/active_storage/sfx/2568/2568-preview.mp3",
                  );
                  audio.play().catch((e) => console.log("Audio play blocked"));
                  if (navigator.vibrate) navigator.vibrate(200);
                  if (scanType === "Kelas") {
                    setClassScanFailCount(0);
                  }
                } else {
                  if (scanType === "Kelas") {
                    setClassScanFailCount((prev) => prev + 1);
                  }
                }

                setScanResult({
                  success: res.success,
                  message: res.message,
                  status: res.status,
                });
              } catch (e) {
                setScanResult({
                  success: false,
                  message: "Gagal memproses pindaian.",
                });
                if (scanType === "Kelas") {
                  setClassScanFailCount((prev) => prev + 1);
                }
              } finally {
                toggleLoader(false);
              }
            },
            (errorMessage: string) => {
              // Ignore frequent scan errors
            },
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
        html5QrCode
          .stop()
          .catch((err: any) => console.error("Unmount stop error:", err));
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
            <h3 className="text-xl font-black text-gray-900 mb-2">
              {confirmModal.title}
            </h3>
            <p className="text-sm text-gray-500 font-medium mb-8 leading-relaxed">
              {confirmModal.entityName ? (
                <>
                  Data <strong>{confirmModal.entityName}</strong> akan dihapus
                  permanen dari sistem.
                </>
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
                onClick={() => {
                  confirmModal.onConfirm();
                  setConfirmModal(null);
                }}
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

  const MassDownloadConfirmModal = () => {
    if (!showMassDownloadConfirm) return null;
    const filtered = students.filter((s) => s.kelas === filterAdminSiswaClass);
    return (
      <div className="fixed inset-0 z-[300] flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm">
        <motion.div
          initial={{ scale: 0.9, opacity: 0 }}
          animate={{ scale: 1, opacity: 1 }}
          className="bg-white p-8 rounded-[2.5rem] shadow-2xl max-w-sm w-full border border-gray-100"
        >
          <div className="flex flex-col items-center text-center">
            <div className="bg-purple-50 p-4 rounded-3xl text-purple-700 mb-6 scale-110">
              <Download size={32} />
            </div>
            <h3 className="text-xl font-black text-gray-900 mb-2">
              Unduh Kartu Siswa
            </h3>
            <p className="text-sm text-gray-500 font-medium mb-8 leading-relaxed">
              Apakah Anda yakin ingin mengunduh{" "}
              <strong className="text-purple-700">{filtered.length}</strong>{" "}
              kartu siswa kelas{" "}
              <strong className="text-green-800">
                {filterAdminSiswaClass}
              </strong>{" "}
              dalam format ZIP?
              <span className="block text-xs font-bold text-gray-400 mt-2">
                Dibutuhkan beberapa detik untuk membuat gambar berkas PDF kartu
                secara otomatis.
              </span>
            </p>
            <div className="flex gap-3 w-full">
              <button
                onClick={() => setShowMassDownloadConfirm(false)}
                className="flex-1 bg-gray-100 text-gray-400 py-4 rounded-2xl font-black text-xs uppercase tracking-widest hover:bg-gray-200 transition-all"
              >
                Batal
              </button>
              <button
                onClick={() => {
                  setShowMassDownloadConfirm(false);
                  toggleLoader(true);
                  setLoaderText(
                    `Menyiapkan data kartu kelas ${filterAdminSiswaClass}...`,
                  );
                  setMassDownloadSiswa(filtered);
                }}
                className="flex-1 bg-purple-800 text-white py-4 rounded-2xl font-black text-xs uppercase tracking-widest shadow-lg shadow-purple-200 hover:bg-purple-700 transition-all scale-105"
              >
                Unduh (.ZIP)
              </button>
            </div>
          </div>
        </motion.div>
      </div>
    );
  };

  const ResetAbsensiModal = () => {
    if (!resetAbsensiTarget) return null;
    return (
      <div className="fixed inset-0 z-[300] flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm">
        <motion.div
          initial={{ scale: 0.9, opacity: 0 }}
          animate={{ scale: 1, opacity: 1 }}
          className="bg-white p-8 rounded-[2.5rem] shadow-2xl max-w-md w-full border border-gray-100"
        >
          <div className="flex flex-col items-center">
            <div className="bg-red-50 p-4 rounded-3xl text-red-600 mb-6 scale-110">
              <Trash2 size={32} />
            </div>
            <h3 className="text-xl font-black text-gray-900 mb-2 text-center">
              Reset Absensi {resetAbsensiTarget}
            </h3>
            <p className="text-sm text-gray-500 font-medium mb-6 text-center leading-relaxed">
              Silakan pilih cakupan penghapusan data absensi{" "}
              <strong>
                {resetAbsensiTarget === "Siswa" ? "Siswa" : "Guru"}
              </strong>
              . Anda dapat memilih bulan & tahun tertentu, atau menghapus
              seluruhnya.
            </p>

            <div className="w-full space-y-4 mb-8">
              <div>
                <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest ml-1">
                  Pilih Bulan & Tahun
                </label>
                <input
                  type="month"
                  value={resetAbsensiBulan}
                  onChange={(e) => setResetAbsensiBulan(e.target.value)}
                  className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-semibold text-sm mt-1 focus:ring-green-600 focus:border-green-600"
                />
              </div>
            </div>

            <div className="flex flex-col gap-2 w-full">
              <button
                type="button"
                disabled={!resetAbsensiBulan}
                onClick={async () => {
                  const target = resetAbsensiTarget;
                  const targetBulan = resetAbsensiBulan;
                  if (!targetBulan) return;
                  setResetAbsensiTarget(null);
                  toggleLoader(true);
                  try {
                    if (target === "Siswa") {
                      await firestoreService.resetAbsensiSiswa(targetBulan);
                      triggerSuccess(
                        "BERHASIL RESET",
                        `Data absensi siswa pada bulan ${targetBulan} berhasil dihapus.`,
                      );
                    } else {
                      await firestoreService.resetAbsensiGuru(targetBulan);
                      triggerSuccess(
                        "BERHASIL RESET",
                        `Data absensi guru pada bulan ${targetBulan} berhasil dihapus.`,
                      );
                    }
                  } catch (err) {
                    alert("Gagal menghapus data.");
                  } finally {
                    toggleLoader(false);
                  }
                }}
                className="w-full bg-red-600 hover:bg-red-700 text-white py-3.5 rounded-2xl font-black text-xs uppercase tracking-widest transition-all disabled:opacity-50 flex items-center justify-center gap-2 shadow-lg shadow-red-150"
              >
                <Trash2 size={14} /> Hapus Hanya Bulan Terpilih (
                {resetAbsensiBulan || "-"})
              </button>

              <button
                type="button"
                onClick={() => {
                  setConfirmModal({
                    show: true,
                    title: `Reset Seluruh Absensi ${resetAbsensiTarget}?`,
                    entityName: `SELURUH DATA ABSENSI ${resetAbsensiTarget.toUpperCase()}`,
                    message: `Aksi ini akan menghapus permanen seluruh data riwayat presensi ${resetAbsensiTarget.toLowerCase()} di database Firestore (tanpa filter bulan). Aksi ini tidak dapat dibatalkan.`,
                    onConfirm: async () => {
                      const target = resetAbsensiTarget;
                      setResetAbsensiTarget(null);
                      toggleLoader(true);
                      try {
                        if (target === "Siswa") {
                          await firestoreService.resetAbsensiSiswa();
                          triggerSuccess(
                            "BERHASIL RESET",
                            "Seluruh data absensi siswa telah berhasil dihapus secara permanen.",
                          );
                        } else {
                          await firestoreService.resetAbsensiGuru();
                          triggerSuccess(
                            "BERHASIL RESET",
                            "Seluruh data absensi guru telah berhasil dihapus secara permanen.",
                          );
                        }
                      } catch (err) {
                        alert("Gagal melakukan reset data.");
                      } finally {
                        toggleLoader(false);
                      }
                    },
                  });
                  setResetAbsensiTarget(null);
                }}
                className="w-full bg-zinc-900 hover:bg-zinc-805 text-white py-3.5 rounded-2xl font-black text-xs uppercase tracking-widest transition-all flex items-center justify-center gap-2"
              >
                <AlertCircle size={14} /> Hapus Seluruh Data (Semua Bulan)
              </button>

              <button
                type="button"
                onClick={() => setResetAbsensiTarget(null)}
                className="w-full bg-zinc-100 hover:bg-zinc-200 text-zinc-500 py-3.5 rounded-2xl font-black text-xs uppercase tracking-widest transition-all text-center mt-2"
              >
                Batalkan
              </button>
            </div>
          </div>
        </motion.div>
      </div>
    );
  };

  const StudentAttendanceDetailModal = () => {
    if (!selectedDetailRecap) return null;

    const { kelas, tanggal, mapel, namaGuru } = selectedDetailRecap;
    const [localStudents, setLocalStudents] = useState<Student[]>([]);
    const [localAttendance, setLocalAttendance] = useState<Attendance[]>([]);
    const [modalLoading, setModalLoading] = useState(true);

    useEffect(() => {
      let isMounted = true;
      const fetchData = async () => {
        setModalLoading(true);
        try {
          // Get students in this class
          const studentsRef = collection(db, "students");
          const studentsQuery = query(studentsRef, where("kelas", "==", kelas));
          const studentsSnap = await getDocs(studentsQuery);
          const sList = studentsSnap.docs.map(
            (doc) => ({ nisn: doc.id, ...doc.data() }) as Student,
          );

          // Get attendance records for this class and date
          const attRef = collection(db, "attendance");
          const attQuery = query(
            attRef,
            where("kelas", "==", kelas),
            where("tanggal", "==", tanggal),
          );
          const attSnap = await getDocs(attQuery);
          const aList = attSnap.docs.map((doc) => doc.data() as Attendance);

          if (isMounted) {
            setLocalStudents(
              sList.sort((a, b) => a.nama.localeCompare(b.nama)),
            );
            setLocalAttendance(aList);
          }
        } catch (error) {
          console.error("Error fetching detail recap data:", error);
        } finally {
          if (isMounted) {
            setModalLoading(false);
          }
        }
      };

      fetchData();
      return () => {
        isMounted = false;
      };
    }, [kelas, tanggal]);

    const countStatus = (status: string) => {
      return localAttendance.filter((a) => a.status === status).length;
    };

    const h = countStatus("Hadir");
    const s = countStatus("Sakit");
    const i = countStatus("Izin");
    const al = countStatus("Alfa");

    return (
      <div className="fixed inset-0 z-[300] flex items-start justify-center p-4 bg-black/60 backdrop-blur-sm overflow-y-auto">
        <motion.div
          initial={{ scale: 0.95, opacity: 0, y: 10 }}
          animate={{ scale: 1, opacity: 1, y: 0 }}
          className="bg-white rounded-[2rem] shadow-2xl max-w-2xl w-full border border-gray-100 overflow-hidden flex flex-col my-auto md:my-10"
        >
          {/* Header */}
          <div className="bg-green-800 p-6 text-white flex justify-between items-start">
            <div className="min-w-0 pr-4">
              <div className="flex items-center gap-2 flex-wrap">
                <span className="bg-white/20 text-white text-[9px] font-black uppercase tracking-widest px-2.5 py-1 rounded-full">
                  Kelas {kelas}
                </span>
                <span className="bg-green-700 text-green-100 text-[9px] font-black uppercase tracking-widest px-2.5 py-1 rounded-full">
                  {mapel}
                </span>
              </div>
              <h3 className="text-xl font-bold mt-2 truncate">
                {formatIndoDate(tanggal)}
              </h3>
              <p className="text-green-200 text-xs mt-1 font-medium truncate">
                Guru Pengajar:{" "}
                <span className="font-extrabold text-white">{namaGuru}</span>
              </p>
            </div>
            <button
              onClick={() => setSelectedDetailRecap(null)}
              className="text-white/80 hover:text-white p-2 hover:bg-white/10 rounded-full transition-all cursor-pointer shrink-0"
            >
              <X size={20} />
            </button>
          </div>

          {/* Stats Bar */}
          <div className="bg-gray-50 border-b border-gray-100 p-4 px-6 grid grid-cols-4 gap-3 text-center">
            <div className="bg-green-50 border border-green-150 rounded-2xl p-2.5">
              <p className="text-[10px] font-black text-green-700 uppercase tracking-widest">
                Hadir
              </p>
              <p className="text-lg font-black text-green-950 mt-0.5">
                {modalLoading ? "..." : h}
              </p>
            </div>
            <div className="bg-orange-50 border border-orange-150 rounded-2xl p-2.5">
              <p className="text-[10px] font-black text-orange-700 uppercase tracking-widest">
                Sakit
              </p>
              <p className="text-lg font-black text-orange-950 mt-0.5">
                {modalLoading ? "..." : s}
              </p>
            </div>
            <div className="bg-blue-50 border border-blue-150 rounded-2xl p-2.5">
              <p className="text-[10px] font-black text-blue-700 uppercase tracking-widest">
                Izin
              </p>
              <p className="text-lg font-black text-blue-950 mt-0.5">
                {modalLoading ? "..." : i}
              </p>
            </div>
            <div className="bg-red-50 border border-red-150 rounded-2xl p-2.5">
              <p className="text-[10px] font-black text-red-700 uppercase tracking-widest">
                Alfa
              </p>
              <p className="text-lg font-black text-red-950 mt-0.5">
                {modalLoading ? "..." : al}
              </p>
            </div>
          </div>

          {/* Download Buttons Bar */}
          <div className="bg-white border-b border-gray-100 p-4 px-6 flex flex-col sm:flex-row sm:items-center justify-between gap-3 shrink-0">
            <span className="text-[10px] font-black text-gray-400 uppercase tracking-widest">
              Unduh Laporan:
            </span>
            <div className="flex gap-2 w-full sm:w-auto">
              <button
                disabled={modalLoading}
                onClick={() => {
                  const dataForExcel = localStudents.map((student, idx) => {
                    const record = localAttendance.find(
                      (at) => at.nisn === student.nisn,
                    );
                    return {
                      No: idx + 1,
                      NISN: student.nisn,
                      "Nama Siswa": student.nama,
                      Kehadiran: record?.status || "Belum Absen",
                      Keterangan: record?.keterangan || "-",
                    };
                  });
                  const ws = XLSX.utils.json_to_sheet(dataForExcel);
                  const wb = XLSX.utils.book_new();
                  XLSX.utils.book_append_sheet(wb, ws, "Presensi Siswa");
                  XLSX.writeFile(
                    wb,
                    `Presensi_Siswa_Kelas_${kelas}_${tanggal}.xlsx`,
                  );
                }}
                className="flex-1 sm:flex-none flex items-center justify-center gap-1.5 bg-green-50 text-green-700 hover:bg-green-100 disabled:opacity-50 px-3.5 py-2.5 rounded-xl text-[10px] font-black uppercase tracking-wider transition-all border border-green-150 cursor-pointer"
              >
                <FileSpreadsheet size={14} /> Excel
              </button>
              <button
                disabled={modalLoading}
                onClick={() => {
                  const doc = new jsPDF("p", "mm", "a4");
                  doc.setFontSize(14);
                  doc.text(`DAFTAR PRESENSI SISWA KELAS ${kelas}`, 14, 20);
                  doc.setFontSize(10);
                  doc.text(`Mata Pelajaran: ${mapel}`, 14, 26);
                  doc.text(`Tanggal: ${formatIndoDate(tanggal)}`, 14, 31);
                  doc.text(`Guru Pengajar: ${namaGuru}`, 14, 36);
                  doc.text(
                    `Rekapitulasi: Hadir (${h}), Sakit (${s}), Izin (${i}), Alfa (${al})`,
                    14,
                    41,
                  );

                  const tableData = localStudents.map((student, idx) => {
                    const record = localAttendance.find(
                      (at) => at.nisn === student.nisn,
                    );
                    return [
                      idx + 1,
                      student.nisn,
                      student.nama,
                      record?.status || "Belum Absen",
                      record?.keterangan || "-",
                    ];
                  });

                  autoTable(doc, {
                    startY: 46,
                    head: [
                      ["No", "NISN", "Nama Siswa", "Kehadiran", "Keterangan"],
                    ],
                    body: tableData,
                    theme: "striped",
                    headStyles: { fillColor: [22, 101, 52] },
                  });
                  doc.save(`Presensi_Siswa_Kelas_${kelas}_${tanggal}.pdf`);
                }}
                className="flex-1 sm:flex-none flex items-center justify-center gap-1.5 bg-red-50 text-red-700 hover:bg-red-100 disabled:opacity-50 px-3.5 py-2.5 rounded-xl text-[10px] font-black uppercase tracking-wider transition-all border border-red-150 cursor-pointer"
              >
                <Download size={14} /> PDF
              </button>
            </div>
          </div>

          {/* Table Container */}
          <div className="p-6 overflow-y-auto max-h-[350px]">
            {modalLoading ? (
              <div className="flex flex-col items-center justify-center py-12 gap-3">
                <div className="w-8 h-8 border-4 border-green-800 border-t-transparent rounded-full animate-spin"></div>
                <p className="text-xs font-bold text-gray-400 uppercase tracking-widest">
                  Memuat data presensi...
                </p>
              </div>
            ) : (
              <table className="w-full text-left text-sm">
                <thead className="text-gray-400 uppercase text-[9px] font-black tracking-widest border-b border-gray-150">
                  <tr>
                    <th className="pb-3 text-center w-12">No</th>
                    <th className="pb-3 pl-2">Siswa</th>
                    <th className="pb-3 text-center w-28">Status</th>
                    <th className="pb-3 pl-4">Keterangan</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-50">
                  {localStudents.map((student, idx) => {
                    const record = localAttendance.find(
                      (at) => at.nisn === student.nisn,
                    );
                    const status = record?.status;

                    let badgeClass = "bg-gray-100 text-gray-400";
                    if (status === "Hadir")
                      badgeClass = "bg-green-100 text-green-700";
                    else if (status === "Sakit")
                      badgeClass = "bg-orange-100 text-orange-700";
                    else if (status === "Izin")
                      badgeClass = "bg-blue-100 text-blue-700";
                    else if (status === "Alfa")
                      badgeClass = "bg-red-100 text-red-700";

                    return (
                      <tr
                        key={student.nisn}
                        className="hover:bg-gray-50/50 transition-colors"
                      >
                        <td className="py-3 text-center text-gray-400 font-bold text-xs">
                          {idx + 1}
                        </td>
                        <td className="py-3 pl-2">
                          <p className="font-bold text-zinc-800 text-xs">
                            {student.nama}
                          </p>
                          <p className="text-[9px] text-gray-400 font-mono">
                            {student.nisn}
                          </p>
                        </td>
                        <td className="py-3 text-center">
                          <span
                            className={`px-2.5 py-1 rounded-full text-[9px] font-black uppercase tracking-wider ${badgeClass}`}
                          >
                            {status || "Belum Absen"}
                          </span>
                        </td>
                        <td
                          className="py-3 pl-4 text-[11px] text-gray-500 italic max-w-[150px] truncate"
                          title={record?.keterangan || ""}
                        >
                          {record?.keterangan || "-"}
                        </td>
                      </tr>
                    );
                  })}
                  {localStudents.length === 0 && (
                    <tr>
                      <td
                        colSpan={4}
                        className="text-center py-8 text-gray-400 font-semibold italic text-xs"
                      >
                        Tidak ada siswa terdaftar di kelas "Kelas {kelas}".
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            )}
          </div>

          {/* Footer */}
          <div className="bg-gray-50 p-4 border-t border-gray-100 flex justify-end">
            <button
              type="button"
              onClick={() => setSelectedDetailRecap(null)}
              className="bg-zinc-800 hover:bg-zinc-900 text-white px-6 py-2.5 rounded-xl text-xs font-black uppercase tracking-widest transition-all shadow-md cursor-pointer"
            >
              Tutup
            </button>
          </div>
        </motion.div>
      </div>
    );
  };

  const MonthlyMapelAttendanceModal = () => {
    if (!showMonthlyMapelAttendance) return null;

    const kelas = rekapMapelFilter.kelas;
    const mapel = rekapMapelFilter.mapel;
    const bulan = rekapMapelFilter.bulan;

    const [modalStudents, setModalStudents] = useState<Student[]>([]);
    const [modalAttendance, setModalAttendance] = useState<Attendance[]>([]);
    const [modalLoading, setModalLoading] = useState(true);

    // Get unique lesson dates sorted chronologically
    const lessonDates = useMemo(() => {
      if (!bulan || !kelas || !mapel) return [];
      const [yearStr, monthStr] = bulan.split("-");
      const yearNum = parseInt(yearStr);
      const monthNum = parseInt(monthStr);
      if (isNaN(yearNum) || isNaN(monthNum)) return [];

      const daysInMonth = new Date(yearNum, monthNum, 0).getDate();
      const monthDates: string[] = [];
      for (let d = 1; d <= daysInMonth; d++) {
        const dayString = d < 10 ? `0${d}` : `${d}`;
        monthDates.push(`${bulan}-${dayString}`);
      }

      const targetSelectedNip = rekapMapelFilter.nip || (session?.role !== "Admin" ? session?.uid : "");
      const matchingSchedules = teachingSchedules.filter((s) => {
        const matchesNip = targetSelectedNip ? s.nip === targetSelectedNip : true;
        const matchesMapel = mapel ? s.mapel === mapel : true;
        const matchesKelas = kelas ? s.kelas === kelas : true;
        return matchesNip && matchesMapel && matchesKelas;
      });

      const datesWithLessons: string[] = [];
      monthDates.forEach((dateStr) => {
        const indonesianDayName = getIndonesianDay(dateStr);
        if (indonesianDayName === "Minggu") return;

        // Check if holiday
        const isHoliday = holidays.some((h) => h.tanggal === dateStr);
        if (isHoliday) return;

        // Check active schedule
        const activeSchedulesForDay = matchingSchedules.filter(
          (s) => s.hari === indonesianDayName
        );

        if (activeSchedulesForDay.length > 0) {
          datesWithLessons.push(dateStr);
        }
      });

      return datesWithLessons.sort();
    }, [bulan, kelas, mapel, rekapMapelFilter.nip, teachingSchedules, holidays, session]);

    const resolvedTeacherName = useMemo(() => {
      const targetSelectedNip = rekapMapelFilter.nip || (session?.role !== "Admin" ? session?.uid : "");
      const sch = teachingSchedules.find(s => {
        const matchesNip = targetSelectedNip ? s.nip === targetSelectedNip : true;
        const matchesMapel = mapel ? s.mapel === mapel : true;
        const matchesKelas = kelas ? s.kelas === kelas : true;
        return matchesNip && matchesMapel && matchesKelas;
      });
      if (sch) return sch.namaGuru;
      if (session?.role !== "Admin") return session?.name || "-";
      return "-";
    }, [rekapMapelFilter.nip, kelas, mapel, teachingSchedules, session]);

    useEffect(() => {
      let isMounted = true;
      const fetchData = async () => {
        if (!kelas || !bulan) {
          setModalLoading(false);
          return;
        }
        setModalLoading(true);
        try {
          // Fetch class students
          const studentsRef = collection(db, "students");
          const studentsQuery = query(studentsRef, where("kelas", "==", kelas));
          const studentsSnap = await getDocs(studentsQuery);
          const sList = studentsSnap.docs.map(
            (doc) => ({ nisn: doc.id, ...doc.data() }) as Student,
          );

          // Fetch attendance records of the selected month
          const attRef = collection(db, "attendance");
          const attQuery = query(
            attRef,
            where("kelas", "==", kelas),
            where("tanggal", ">=", `${bulan}-01`),
            where("tanggal", "<=", `${bulan}-31`),
          );
          const attSnap = await getDocs(attQuery);
          const aList = attSnap.docs.map((doc) => doc.data() as Attendance);

          if (isMounted) {
            setModalStudents(
              sList.sort((a, b) => a.nama.localeCompare(b.nama)),
            );
            setModalAttendance(aList);
          }
        } catch (error) {
          console.error("Error fetching monthly mapel attendance data:", error);
        } finally {
          if (isMounted) {
            setModalLoading(false);
          }
        }
      };

      fetchData();
      return () => {
        isMounted = false;
      };
    }, [kelas, bulan]);

    // Handle Excel download for monthly mapel recap
    const handleExportExcel = () => {
      if (modalStudents.length === 0) return;
      const exportData = modalStudents.map((siswa, idx) => {
        const rowData: any = {
          No: idx + 1,
          Nama: siswa.nama,
          NISN: siswa.nisn,
        };

        let hCount = 0;
        let sCount = 0;
        let iCount = 0;
        let aCount = 0;

        lessonDates.forEach(date => {
          const rec = modalAttendance.find(
            (la) => la.nisn === siswa.nisn && la.tanggal === date
          );
          const dayNum = date.split("-")[2];
          if (rec) {
            rowData[`Tanggal ${dayNum}`] = rec.status === "Hadir" ? "✓" : rec.status.charAt(0);
            if (rec.status === "Hadir") hCount++;
            else if (rec.status === "Sakit") sCount++;
            else if (rec.status === "Izin") iCount++;
            else if (rec.status === "Alfa") aCount++;
          } else {
            rowData[`Tanggal ${dayNum}`] = "-";
          }
        });

        rowData["Hadir"] = hCount;
        rowData["Sakit"] = sCount;
        rowData["Izin"] = iCount;
        rowData["Alfa"] = aCount;
        rowData["Kehadiran (%)"] = lessonDates.length > 0
          ? `${Math.round((hCount / lessonDates.length) * 100)}%`
          : "0%";

        return rowData;
      });

      const ws = XLSX.utils.json_to_sheet(exportData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Siswa");

      XLSX.writeFile(
        wb,
        `Rekap_Bulanan_${mapel}_Kls_${kelas}_${bulan}.xlsx`,
      );
    };

    // Handle PDF download
    const handleExportPDF = () => {
      if (modalStudents.length === 0) return;
      const doc = new jsPDF("l", "mm", "a4");
      doc.setFontSize(14);
      doc.text("REKAPITULASI KEHADIRAN SISWA BULANAN", 14, 15);
      doc.setFontSize(10);
      doc.text(`Guru Pengajar : ${resolvedTeacherName || "-"}`, 14, 21);
      doc.text(`Mata Pelajaran : ${mapel} | Kelas: ${kelas}`, 14, 26);
      doc.text(`Periode Bulan  : ${getMonthYearText(bulan)}`, 14, 31);

      const headers = [
        "No",
        "Nama Siswa",
        "NISN",
        ...lessonDates.map(d => d.split("-")[2]),
        "H",
        "S",
        "I",
        "A",
        "%"
      ];

      const rows = modalStudents.map((siswa, idx) => {
        let hCount = 0;
        let sCount = 0;
        let iCount = 0;
        let aCount = 0;

        const dayStatuses = lessonDates.map(date => {
          const rec = modalAttendance.find(
            (la) => la.nisn === siswa.nisn && la.tanggal === date
          );
          if (rec) {
            if (rec.status === "Hadir") { hCount++; return "✓"; }
            if (rec.status === "Sakit") { sCount++; return "S"; }
            if (rec.status === "Izin") { iCount++; return "I"; }
            if (rec.status === "Alfa") { aCount++; return "A"; }
          }
          return "-";
        });

        const pct = lessonDates.length > 0
          ? `${Math.round((hCount / lessonDates.length) * 100)}%`
          : "0%";

        return [
          idx + 1,
          siswa.nama,
          siswa.nisn,
          ...dayStatuses,
          hCount,
          sCount,
          iCount,
          aCount,
          pct
        ];
      });

      autoTable(doc, {
        startY: 37,
        head: [headers],
        body: rows,
        theme: "grid",
        styles: { fontSize: 7, cellPadding: 1 },
        columnStyles: {
          1: { halign: "left" }
        }
      });

      doc.save(`Rekap_Bulanan_${mapel}_Kls_${kelas}_${bulan}.pdf`);
    };

    return (
      <div className="fixed inset-0 z-[300] flex items-start justify-center p-4 bg-black/60 backdrop-blur-sm overflow-y-auto">
        <motion.div
          initial={{ scale: 0.95, opacity: 0, y: 10 }}
          animate={{ scale: 1, opacity: 1, y: 0 }}
          className="bg-white rounded-[2rem] shadow-2xl max-w-5xl w-full border border-gray-100 overflow-hidden flex flex-col my-auto md:my-10"
        >
          {/* Header */}
          <div className="bg-green-800 p-6 text-white flex justify-between items-start">
            <div className="min-w-0 pr-4">
              <div className="flex items-center gap-2 flex-wrap">
                <span className="bg-white/20 text-white text-[9px] font-black uppercase tracking-widest px-2.5 py-1 rounded-full">
                  Kelas {kelas || "-"}
                </span>
                <span className="bg-green-700 text-green-100 text-[9px] font-black uppercase tracking-widest px-2.5 py-1 rounded-full">
                  {mapel || "-"}
                </span>
                <span className="bg-zinc-900/40 text-green-100 text-[9px] font-black uppercase tracking-widest px-2.5 py-1 rounded-full text-nowrap">
                  {getMonthYearText(bulan)}
                </span>
              </div>
              <h3 className="text-xl font-bold mt-2 truncate">
                Rekap Absensi Bulanan Siswa
              </h3>
              <p className="text-green-200 text-xs mt-1 font-medium truncate">
                Guru Pengajar:{" "}
                <span className="font-extrabold text-white">{resolvedTeacherName || "-"}</span>
              </p>
            </div>
            <button
              onClick={() => setShowMonthlyMapelAttendance(false)}
              className="text-white/80 hover:text-white p-2 hover:bg-white/10 rounded-full transition-all cursor-pointer shrink-0"
            >
              <X size={20} />
            </button>
          </div>

          {/* Action Button Bar */}
          <div className="bg-gray-50 border-b border-gray-100 px-6 py-4 flex flex-col sm:flex-row justify-between items-center gap-3">
            <div className="text-xs font-bold text-gray-500">
              {modalLoading ? (
                <span>Mengambil data...</span>
              ) : (
                <span>Total: <strong className="text-zinc-800">{modalStudents.length} Siswa</strong> | <strong className="text-zinc-800">{lessonDates.length} Hari Belajar</strong></span>
              )}
            </div>
            <div className="flex items-center gap-2 w-full sm:w-auto">
              <button
                onClick={handleExportExcel}
                disabled={modalLoading || modalStudents.length === 0}
                className="flex-1 sm:flex-none flex items-center justify-center gap-1.5 bg-green-50 text-green-700 hover:bg-green-100 disabled:opacity-50 px-3.5 py-2 rounded-xl text-[10px] font-black uppercase tracking-wider transition-all border border-green-150 cursor-pointer"
              >
                <FileSpreadsheet size={14} /> Excel
              </button>
              <button
                onClick={handleExportPDF}
                disabled={modalLoading || modalStudents.length === 0}
                className="flex-1 sm:flex-none flex items-center justify-center gap-1.5 bg-red-50 text-red-700 hover:bg-red-100 disabled:opacity-50 px-3.5 py-2 rounded-xl text-[10px] font-black uppercase tracking-wider transition-all border border-red-150 cursor-pointer"
              >
                <FileText size={14} /> PDF
              </button>
            </div>
          </div>

          {/* Modal content */}
          <div className="p-6 overflow-y-auto max-h-[60vh]">
            {modalLoading ? (
              <div className="py-20 text-center text-zinc-400 font-bold">
                Dalam proses memuat database...
              </div>
            ) : modalStudents.length === 0 ? (
              <div className="py-20 text-center text-zinc-400 font-bold">
                Tidak ada data siswa ditemukan untuk kelas ini.
              </div>
            ) : (
              <div className="overflow-x-auto rounded-3xl border border-gray-100 shadow-sm">
                <table className="w-full text-left text-sm min-w-[700px]">
                  <thead className="bg-gray-50 text-gray-400 uppercase text-[10px] font-black tracking-widest border-b border-gray-100">
                    <tr>
                      <th className="px-4 py-3 text-center">No</th>
                      <th className="px-4 py-3">Nama Siswa</th>
                      <th className="px-4 py-3">NISN</th>
                      {lessonDates.map(date => {
                        const dayNum = date.split("-")[2];
                        return (
                          <th key={date} className="px-2 py-3 text-center min-w-[40px]" title={date}>
                            {dayNum}
                          </th>
                        );
                      })}
                      <th className="px-3 py-3 text-center bg-green-50/55 text-green-800">H</th>
                      <th className="px-3 py-3 text-center bg-orange-50/55 text-orange-850">S</th>
                      <th className="px-3 py-3 text-center bg-blue-50/55 text-blue-800">I</th>
                      <th className="px-3 py-3 text-center bg-red-50/55 text-red-800">A</th>
                      <th className="px-3 py-3 text-center bg-zinc-100 text-zinc-800">%</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-100">
                    {modalStudents.map((siswa, idx) => {
                      let hCount = 0;
                      let sCount = 0;
                      let iCount = 0;
                      let aCount = 0;

                      const dateCells = lessonDates.map(date => {
                        const rec = modalAttendance.find(
                          (la) => la.nisn === siswa.nisn && la.tanggal === date
                        );
                        let statusChar = "-";
                        let colorClass = "text-gray-300";

                        if (rec) {
                          if (rec.status === "Hadir") {
                            hCount++;
                            statusChar = "✓";
                            colorClass = "text-green-600 font-black";
                          } else if (rec.status === "Sakit") {
                            sCount++;
                            statusChar = "S";
                            colorClass = "text-orange-600 font-black";
                          } else if (rec.status === "Izin") {
                            iCount++;
                            statusChar = "I";
                            colorClass = "text-blue-600 font-black";
                          } else if (rec.status === "Alfa") {
                            aCount++;
                            statusChar = "A";
                            colorClass = "text-red-600 font-black";
                          }
                        }

                        return (
                          <td key={date} className={`px-2 py-3 text-center ${colorClass}`}>
                            {statusChar}
                          </td>
                        );
                      });

                      const pct = lessonDates.length > 0
                        ? Math.round((hCount / lessonDates.length) * 100)
                        : 0;

                      return (
                        <tr key={siswa.nisn} className="hover:bg-zinc-50 transition-colors">
                          <td className="px-4 py-3 text-center font-bold text-gray-400 text-xs">
                            {idx + 1}
                          </td>
                          <td className="px-4 py-3 font-bold text-zinc-800 text-xs text-nowrap">
                            {siswa.nama}
                          </td>
                          <td className="px-4 py-3 text-xs font-mono text-gray-400">
                            {siswa.nisn}
                          </td>
                          {dateCells}
                          <td className="px-3 py-3 text-center font-bold bg-green-50/30 text-green-700 text-xs">
                            {hCount}
                          </td>
                          <td className="px-3 py-3 text-center font-bold bg-orange-50/30 text-orange-700 text-xs">
                            {sCount}
                          </td>
                          <td className="px-3 py-3 text-center font-bold bg-blue-50/30 text-blue-700 text-xs">
                            {iCount}
                          </td>
                          <td className="px-3 py-3 text-center font-bold bg-red-50/30 text-red-700 text-xs">
                            {aCount}
                          </td>
                          <td className="px-3 py-3 text-center font-black bg-zinc-50 text-green-800 text-xs text-nowrap">
                            {pct}%
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        </motion.div>
      </div>
    );
  };

  const [showSiswaModal, setShowSiswaModal] = useState(false);
  const [editingSiswa, setEditingSiswa] = useState<Student | null>(null);

  // States for Naik Kelas (Promotion) feature
  const [showNaikKelasModal, setShowNaikKelasModal] = useState(false);
  const [naikKelasOrigin, setNaikKelasOrigin] = useState("");
  const [naikKelasTarget, setNaikKelasTarget] = useState("");
  const [naikKelasSelectedStudents, setNaikKelasSelectedStudents] = useState<
    string[]
  >([]);
  const [naikKelasSearchKeyword, setNaikKelasSearchKeyword] = useState("");

  // Merge database classrooms with specified items to cover all eventualities gracefully
  const naikKelasOriginOptions = useMemo(() => {
    const specified = [
      "VII A",
      "VII B",
      "VII C",
      "VII D",
      "VII E",
      "VII F",
      "VIII A",
      "VIII B",
      "VIII C",
      "VIII D",
      "VIII E",
      "VIII F",
      "IX A",
      "IX B",
      "IX C",
      "IX D",
      "IX E",
      "IX F",
    ];
    // Filter database classrooms that are not already in specified
    const dbOthers = (classrooms || [])
      .map((c) => c.nama)
      .filter(
        (name) =>
          !specified.includes(name) &&
          (name.startsWith("VII") ||
            name.startsWith("7") ||
            name.startsWith("VIII") ||
            name.startsWith("8") ||
            name.startsWith("IX") ||
            name.startsWith("9")),
      );
    return [...specified, ...dbOthers];
  }, [classrooms]);

  const naikKelasTargetOptions = useMemo(() => {
    if (
      naikKelasOrigin &&
      (naikKelasOrigin.startsWith("IX") || naikKelasOrigin.startsWith("9"))
    ) {
      return ["Tamat"];
    }
    const specified = [
      "VIII A",
      "VIII B",
      "VIII C",
      "VIII D",
      "VIII E",
      "VIII F",
      "IX A",
      "IX B",
      "IX C",
      "IX D",
      "IX E",
      "IX F",
    ];
    const dbOthers = (classrooms || [])
      .map((c) => c.nama)
      .filter(
        (name) =>
          !specified.includes(name) &&
          (name.startsWith("VIII") ||
            name.startsWith("8") ||
            name.startsWith("IX") ||
            name.startsWith("9")),
      );
    return [...specified, ...dbOthers];
  }, [classrooms, naikKelasOrigin]);

  const naikKelasFilteredStudents = useMemo(() => {
    if (!naikKelasOrigin) return [];
    return students.filter((s) => {
      const matchClass = s.kelas === naikKelasOrigin;
      if (!matchClass) return false;
      if (!naikKelasSearchKeyword) return true;
      const sLower = naikKelasSearchKeyword.toLowerCase();
      return s.nama.toLowerCase().includes(sLower) || s.nisn.includes(sLower);
    });
  }, [students, naikKelasOrigin, naikKelasSearchKeyword]);

  // Handle select/unselect all shown
  const handleToggleSelectAllNaikKelas = () => {
    const allFilteredNisns = naikKelasFilteredStudents.map((s) => s.nisn);
    const areAllSelected = allFilteredNisns.every((nisn) =>
      naikKelasSelectedStudents.includes(nisn),
    );
    if (areAllSelected) {
      // Unselect everyone currently filtered
      setNaikKelasSelectedStudents((prev) =>
        prev.filter((nisn) => !allFilteredNisns.includes(nisn)),
      );
    } else {
      // Select everyone currently filtered
      setNaikKelasSelectedStudents((prev) => {
        const union = new Set([...prev, ...allFilteredNisns]);
        return Array.from(union);
      });
    }
  };

  const handleToggleSelectStudentNaikKelas = (nisn: string) => {
    setNaikKelasSelectedStudents((prev) =>
      prev.includes(nisn) ? prev.filter((n) => n !== nisn) : [...prev, nisn],
    );
  };

  // Pre-select all students of the chosen class by default
  useEffect(() => {
    if (showNaikKelasModal && naikKelasOrigin) {
      const allNisns = students
        .filter((s) => s.kelas === naikKelasOrigin)
        .map((s) => s.nisn);
      setNaikKelasSelectedStudents(allNisns);
    } else {
      setNaikKelasSelectedStudents([]);
    }
  }, [naikKelasOrigin, showNaikKelasModal, students]);

  // Autoselect 'Tamat' when Class IX is chosen as Origin Class
  useEffect(() => {
    if (
      naikKelasOrigin &&
      (naikKelasOrigin.startsWith("IX") || naikKelasOrigin.startsWith("9"))
    ) {
      setNaikKelasTarget("Tamat");
    } else {
      if (naikKelasTarget === "Tamat") {
        setNaikKelasTarget("");
      }
    }
  }, [naikKelasOrigin]);

  const handleProsesNaikKelas = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!naikKelasOrigin || !naikKelasTarget) {
      alert("Harap pilih kelas asal dan kelas tujuan.");
      return;
    }
    if (naikKelasSelectedStudents.length === 0) {
      alert("Harap pilih setidaknya satu siswa untuk dinaikkan kelas.");
      return;
    }

    const isTamat =
      naikKelasTarget === "Tamat" ||
      naikKelasOrigin.startsWith("IX") ||
      naikKelasOrigin.startsWith("9");

    setConfirmModal({
      show: true,
      title: isTamat
        ? "Konfirmasi Kelulusan (Tamat)?"
        : "Konfirmasi Naik Kelas?",
      message: isTamat
        ? `Apakah anda yakin ingin memproses kelulusan? Sebanyak ${naikKelasSelectedStudents.length} siswa dari kelas ${naikKelasOrigin} yang terpilih akan DIHAPUS PERMANEN dari database karena telah dinyatakan lulus dari madrasah.`
        : `Sebanyak ${naikKelasSelectedStudents.length} siswa akan diubah kelasnya dari ${naikKelasOrigin} menjadi ${naikKelasTarget}. Aksi ini akan langsung disinkronkan ke Firestore.`,
      entityName: `${naikKelasSelectedStudents.length} Siswa`,
      onConfirm: async () => {
        toggleLoader(true);
        try {
          // Filter students who are selected
          const studentsToUpdate = students.filter((s) =>
            naikKelasSelectedStudents.includes(s.nisn),
          );
          if (isTamat) {
            for (const s of studentsToUpdate) {
              await firestoreService.hapusSiswa(s.nisn);
            }
            triggerSuccess(
              "KELULUSAN BERHASIL",
              `Sukses! Sebanyak ${studentsToUpdate.length} siswa kelas IX telah dinyatakan lulus & data dihapus dari database.`,
            );
          } else {
            for (const s of studentsToUpdate) {
              await firestoreService.saveSiswa({
                ...s,
                kelas: naikKelasTarget,
              });
            }
            triggerSuccess(
              "BERHASIL",
              `Hore! Sebanyak ${studentsToUpdate.length} siswa berhasil dinaikkan kelas ke ${naikKelasTarget}.`,
            );
          }
          setShowNaikKelasModal(false);
          setNaikKelasOrigin("");
          setNaikKelasTarget("");
          setNaikKelasSelectedStudents([]);
          setNaikKelasSearchKeyword("");
        } catch (e) {
          console.error(e);
          alert(
            isTamat
              ? "Gagal memproses kelulusan/penghapusan siswa kelas IX."
              : "Gagal melakukan proses naik kelas.",
          );
        } finally {
          toggleLoader(false);
        }
      },
    });
  };
  const [showGuruModal, setShowGuruModal] = useState(false);
  const [editingGuru, setEditingGuru] = useState<any | null>(null);
  const [showKelasModal, setShowKelasModal] = useState(false);
  const [editingKelas, setEditingKelas] = useState<any | null>(null);
  const [showSuccessToast, setShowSuccessToast] = useState<{
    title: string;
    message: string;
  } | null>(null);
  const [showWarningToast, setShowWarningToast] = useState<{
    title: string;
    message: string;
  } | null>(null);
  const [newLibur, setNewLibur] = useState({ tanggal: "", keterangan: "" });
  const [selectedCalendarMonth, setSelectedCalendarMonth] = useState(() =>
    new Date().toISOString().slice(0, 7),
  );

  const [showCelebration, setShowCelebration] = useState(false);

  const triggerSuccess = (title: string, message: string) => {
    setShowSuccessToast({ title, message });
    setShowCelebration(true);
    setSearchTermGuru("");
    setSearchTermSiswa("");
    setSearchTermKelas("");
    setSearchTermJadwal("");
    setPagination({
      guru: 0,
      siswa: 0,
      kelas: 0,
      jadwal: 0,
      absensiSiswa: 0,
      absensiGuru: 0,
      rekapMapel: 0,
    });
    setTimeout(() => {
      setShowSuccessToast(null);
      setShowCelebration(false);
    }, 3000);
  };

  const triggerWarning = (title: string, message: string) => {
    setShowWarningToast({ title, message });
    setTimeout(() => {
      setShowWarningToast(null);
    }, 4000);
  };

  const triggerError = (title: string, message: string) => {
    alert(`${title}: ${message}`);
  };
  const [mobileExtraMenuOpen, setMobileExtraMenuOpen] = useState(false);
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);
  const [showJadwalModal, setShowJadwalModal] = useState(false);
  const [editingJadwal, setEditingJadwal] = useState<TeachingSchedule | null>(
    null,
  );
  const [showSessionDetail, setShowSessionDetail] = useState<{
    name: string;
    mapping: any[];
  } | null>(null);

  const printBarcode = (className: string) => {
    const win = window.open("", "_blank");
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
      await new Promise((resolve) => setTimeout(resolve, 1500));

      // Force reload images with crossOrigin for CORS compatibility
      const images = Array.from(el.querySelectorAll("img"));
      await Promise.all(
        images.map((img) => {
          if (img.complete && img.naturalWidth !== 0) return Promise.resolve();
          return new Promise((resolve) => {
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
        }),
      );

      // Additional small delay for layout stabilization
      await new Promise((resolve) => setTimeout(resolve, 300));

      // Try html-to-image first
      try {
        const dataUrl = await htmlToImage.toPng(el, {
          pixelRatio: 2,
          backgroundColor: "#ffffff",
          cacheBust: true,
          style: {
            transform: "none",
            visibility: "visible",
            display: "flex",
            opacity: "1",
            margin: "0",
            padding: "0",
          },
        });

        if (dataUrl && dataUrl.length > 500) {
          console.log("Capture with html-to-image successful.");
          return dataUrl;
        }
      } catch (innerErr) {
        console.warn("html-to-image failed, trying html2canvas...", innerErr);
      }

      // html2canvas fallback
      const html2canvas = (await import("html2canvas")).default;
      const canvas = await html2canvas(el, {
        useCORS: true,
        allowTaint: true,
        backgroundColor: "#ffffff",
        scale: 2,
        logging: false,
        onclone: (clonedDoc) => {
          const win = clonedDoc.defaultView || window;
          const clonedEl = clonedDoc.getElementById(elementId);
          if (clonedEl) {
            clonedEl.style.cssText +=
              "transform: none !important; visibility: visible !important; display: flex !important; opacity: 1 !important; position: static !important; background: white !important;";

            // Recursively remove oklch/oklab from ALL stylesheets in the clone
            // html2canvas parser crashes when it encounters these keywords anywhere
            Array.from(clonedDoc.styleSheets).forEach((sheet) => {
              try {
                const rules = sheet.cssRules;
                for (let i = rules.length - 1; i >= 0; i--) {
                  if (
                    rules[i].cssText.includes("oklch") ||
                    rules[i].cssText.includes("oklab")
                  ) {
                    sheet.deleteRule(i);
                  }
                }
              } catch (e) {
                // Cross-origin sheets might throw, so we fallback to fixing style tags
              }
            });

            const styleTags = clonedDoc.querySelectorAll("style");
            styleTags.forEach((tag) => {
              try {
                tag.innerHTML = tag.innerHTML.replace(
                  /(oklch|oklab)\s*\([^)]+\)/g,
                  "#333333",
                );
              } catch (e) {}
            });

            // Recursively find and fix oklch colors in inline/computed styles for elements
            const allElements = clonedEl.querySelectorAll("*");
            const canvasHelper = clonedDoc.createElement("canvas");
            const ctx = canvasHelper.getContext("2d");

            allElements.forEach((child) => {
              const element = child as HTMLElement;

              // Remove filter/backdropFilter as they cause issues
              element.style.filter = "none";
              element.style.backdropFilter = "none";

              // Properties to check for oklch colors
              const props = [
                "color",
                "backgroundColor",
                "borderColor",
                "outlineColor",
                "fill",
                "stroke",
                "stopColor",
              ];

              props.forEach((prop) => {
                // Try to get value carefully to avoid triggering crash if possible
                let val = "";
                try {
                  val = element.style.getPropertyValue(prop);
                  if (!val || val.includes("oklch") || val.includes("oklab")) {
                    const computed = win.getComputedStyle(element);
                    val = computed.getPropertyValue(prop);
                  }
                } catch (e) {}

                if (val && (val.includes("oklch") || val.includes("oklab"))) {
                  let converted = false;
                  if (ctx) {
                    try {
                      ctx.fillStyle = val;
                      const rgbVal = ctx.fillStyle;
                      if (
                        rgbVal &&
                        !rgbVal.includes("oklch") &&
                        !rgbVal.includes("oklab") &&
                        rgbVal !== "#000000"
                      ) {
                        element.style.setProperty(prop, rgbVal, "important");
                        converted = true;
                      }
                    } catch (e) {}
                  }

                  if (!converted) {
                    if (prop === "color" || prop === "stroke")
                      element.style.setProperty(prop, "#000000", "important");
                    else if (prop === "backgroundColor" || prop === "fill")
                      element.style.setProperty(prop, "#ffffff", "important");
                    else
                      element.style.setProperty(
                        prop,
                        "transparent",
                        "important",
                      );
                  }
                }
              });
            });
          }
        },
      });

      const canvasDataUrl = canvas.toDataURL("image/png");
      console.log("Capture with html2canvas successful.");
      return canvasDataUrl;
    } catch (err) {
      console.error("All capture methods failed:", err);
      try {
        // Absolute last resort: simplest possible capture
        return await htmlToImage.toPng(el);
      } catch (finalErr) {
        alert(
          `Gagal mengambil gambar: ${err instanceof Error ? err.message : "Unknown Error"}`,
        );
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
      await new Promise((resolve) => (img.onload = resolve));

      // Use logical dimensions (divided by 2 because we used pixelRatio: 2)
      const logicalWidth = img.width / 2;
      const logicalHeight = img.height / 2;

      const pdf = new jsPDF({
        orientation: logicalWidth > logicalHeight ? "l" : "p",
        unit: "px",
        format: [logicalWidth, logicalHeight],
      });

      pdf.addImage(dataUrl, "PNG", 0, 0, logicalWidth, logicalHeight);

      pdf.save(fileName.replace(".png", ".pdf"));
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

  const captureBatchElement = async (el: HTMLElement, elementId: string) => {
    const originalStyle = el.style.cssText;

    // Temporarily make it visible and independent for capture
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
      // Delay to ensure the DOM has rendered the new styles and images
      await new Promise((resolve) => setTimeout(resolve, 300));

      // Force reload images inside for CORS compatibility
      const images = Array.from(el.querySelectorAll("img"));
      await Promise.all(
        images.map((img) => {
          if (img.complete && img.naturalWidth !== 0) return Promise.resolve();
          return new Promise((resolve) => {
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
        }),
      );

      // Additional small delay for layout stabilization
      await new Promise((resolve) => setTimeout(resolve, 150));

      // Try html-to-image first
      try {
        const dataUrl = await htmlToImage.toPng(el, {
          pixelRatio: 2,
          backgroundColor: "#ffffff",
          cacheBust: true,
          style: {
            transform: "none",
            visibility: "visible",
            display: "flex",
            opacity: "1",
            margin: "0",
            padding: "0",
          },
        });
        if (dataUrl && dataUrl.length > 500) {
          return dataUrl;
        }
      } catch (innerErr) {
        console.warn(
          "Batch html-to-image failed, trying html2canvas...",
          innerErr,
        );
      }

      // html2canvas fallback
      const html2canvas = (await import("html2canvas")).default;
      const canvas = await html2canvas(el, {
        useCORS: true,
        allowTaint: true,
        backgroundColor: "#ffffff",
        scale: 2,
        logging: false,
        onclone: (clonedDoc) => {
          const win = clonedDoc.defaultView || window;
          const clonedEl = clonedDoc.getElementById(elementId);
          if (clonedEl) {
            clonedEl.style.cssText +=
              "transform: none !important; visibility: visible !important; display: flex !important; opacity: 1 !important; position: static !important; background: white !important;";

            // Recursively remove oklch/oklab from ALL stylesheets in the clone
            Array.from(clonedDoc.styleSheets).forEach((sheet) => {
              try {
                const rules = sheet.cssRules;
                for (let i = rules.length - 1; i >= 0; i--) {
                  if (
                    rules[i].cssText.includes("oklch") ||
                    rules[i].cssText.includes("oklab")
                  ) {
                    sheet.deleteRule(i);
                  }
                }
              } catch (e) {}
            });

            const styleTags = clonedDoc.querySelectorAll("style");
            styleTags.forEach((tag) => {
              try {
                tag.innerHTML = tag.innerHTML.replace(
                  /(oklch|oklab)\s*\([^)]+\)/g,
                  "#333333",
                );
              } catch (e) {}
            });

            // Recursively find and fix oklch colors in inline/computed styles for elements
            const allElements = clonedEl.querySelectorAll("*");
            const canvasHelper = clonedDoc.createElement("canvas");
            const ctx = canvasHelper.getContext("2d");

            allElements.forEach((child) => {
              const element = child as HTMLElement;
              element.style.filter = "none";
              element.style.backdropFilter = "none";

              const props = [
                "color",
                "backgroundColor",
                "borderColor",
                "outlineColor",
                "fill",
                "stroke",
                "stopColor",
              ];
              props.forEach((prop) => {
                let val = "";
                try {
                  val = element.style.getPropertyValue(prop);
                  if (!val || val.includes("oklch") || val.includes("oklab")) {
                    const computed = win.getComputedStyle(element);
                    val = computed.getPropertyValue(prop);
                  }
                } catch (e) {}

                if (val && (val.includes("oklch") || val.includes("oklab"))) {
                  let converted = false;
                  if (ctx) {
                    try {
                      ctx.fillStyle = val;
                      const rgbVal = ctx.fillStyle;
                      if (
                        rgbVal &&
                        !rgbVal.includes("oklch") &&
                        !rgbVal.includes("oklab") &&
                        rgbVal !== "#000000"
                      ) {
                        element.style.setProperty(prop, rgbVal, "important");
                        converted = true;
                      }
                    } catch (e) {}
                  }
                  if (!converted) {
                    if (prop === "color" || prop === "stroke")
                      element.style.setProperty(prop, "#000000", "important");
                    else if (prop === "backgroundColor" || prop === "fill")
                      element.style.setProperty(prop, "#ffffff", "important");
                    else
                      element.style.setProperty(
                        prop,
                        "transparent",
                        "important",
                      );
                  }
                }
              });
            });
          }
        },
      });

      return canvas.toDataURL("image/png");
    } catch (err) {
      console.error("Batch capture failed using any strategy:", err);
      try {
        return await htmlToImage.toPng(el);
      } catch (finalErr) {
        return null;
      }
    } finally {
      el.style.cssText = originalStyle;
    }
  };

  const downloadMassSiswaKartu = async () => {
    if (!filterAdminSiswaClass) {
      alert(
        "Silakan pilih kelas terlebih dahulu pada filter kelas di sebelah kiri untuk mengunduh kartu secara massal.",
      );
      return;
    }

    const filtered = students.filter((s) => s.kelas === filterAdminSiswaClass);
    if (filtered.length === 0) {
      alert(`Tidak ada data siswa untuk kelas ${filterAdminSiswaClass}.`);
      return;
    }

    setShowMassDownloadConfirm(true);
  };

  useEffect(() => {
    if (massDownloadSiswa.length === 0) return;

    const processBatch = async () => {
      try {
        setLoaderText("Memulai rendering kartu...");
        // Wait 3 seconds for standard react render cycle, background templates, images and barcodes to fully generate in the DOM
        await new Promise((resolve) => setTimeout(resolve, 3000));

        const zip = new JSZip();
        let successCount = 0;

        for (let i = 0; i < massDownloadSiswa.length; i++) {
          const student = massDownloadSiswa[i];
          setLoaderText(
            `Membuat kartu [${i + 1}/${massDownloadSiswa.length}]: ${student.nama}...`,
          );

          const elementId = `mass-print-student-card-${student.nisn}`;
          const el = document.getElementById(elementId);
          if (el) {
            const dataUrl = await captureBatchElement(el, elementId);
            if (dataUrl) {
              const img = new Image();
              img.src = dataUrl;
              await new Promise((r) => (img.onload = r));

              const logicalWidth = img.width / 2;
              const logicalHeight = img.height / 2;

              const pdf = new jsPDF({
                orientation: logicalWidth > logicalHeight ? "l" : "p",
                unit: "px",
                format: [logicalWidth, logicalHeight],
              });

              pdf.addImage(dataUrl, "PNG", 0, 0, logicalWidth, logicalHeight);
              const pdfOutput = pdf.output("arraybuffer");

              const safeName = student.nama.replace(/[/\\?%*:|"<>]/g, "-");
              zip.file(`${student.nisn}_${safeName}.pdf`, pdfOutput);
              successCount++;
            } else {
              console.warn(
                `Could not capture card for student ${student.nama}`,
              );
            }
          } else {
            console.warn(`Element #${elementId} not found in DOM`);
          }
        }

        if (successCount > 0) {
          setLoaderText("Mengompresi semua kartu ke dalam file ZIP...");
          const content = await zip.generateAsync({ type: "blob" });

          const url = URL.createObjectURL(content);
          const a = document.createElement("a");
          a.style.display = "none";
          a.href = url;
          a.download = `Kartu-Siswa-Kelas-${filterAdminSiswaClass}.zip`;
          document.body.appendChild(a);
          a.click();
          document.body.removeChild(a);
          URL.revokeObjectURL(url);

          triggerSuccess(
            "BERHASIL UNDUH",
            `Sukses mengunduh ${successCount} kartu siswa kelas ${filterAdminSiswaClass}.`,
          );
        } else {
          alert(
            "Gagal memproses kartu siswa. Elemen tidak terfaktorkan dengan benar.",
          );
        }
      } catch (err) {
        console.error("Batch card download failed:", err);
        alert("Gagal mengunduh kartu massal.");
      } finally {
        setMassDownloadSiswa([]);
        setLoaderText("");
        toggleLoader(false);
      }
    };

    processBatch();
  }, [massDownloadSiswa]);

  useEffect(() => {
    setLoginNisn("");
    setLoginUser("");
    setLoginPass("");
    setLoginError("");
  }, [loginRole]);

  const [siswaAbsensiManual, setSiswaAbsensiManual] = useState({
    nisn: "",
    status: "Hadir",
    ket: "",
    tanggal: new Date().toISOString().split("T")[0],
  });
  const [selectedStudentCard, setSelectedStudentCard] =
    useState<Student | null>(null);

  const exportExcel = () => {
    const isSiswa = rekapFilter.type === "Siswa";
    const data = isSiswa
      ? students
          .filter((s) =>
            rekapFilter.kelas ? s.kelas === rekapFilter.kelas : true,
          )
          .map((s, idx) => {
            const monthAttendance = attendance.filter(
              (a) =>
                a.nisn === s.nisn && a.tanggal.startsWith(rekapFilter.bulan),
            );
            const hadir = monthAttendance.filter(
              (a) => a.status === "Hadir",
            ).length;
            const izin = monthAttendance.filter(
              (a) => a.status === "Izin",
            ).length;
            const sakit = monthAttendance.filter(
              (a) => a.status === "Sakit",
            ).length;
            const alfa = monthAttendance.filter(
              (a) => a.status === "Alfa",
            ).length;
            const totalLambat = monthAttendance.reduce(
              (sum, a) => sum + (a.terlambat || 0),
              0,
            );
            const [y, m] = rekapFilter.bulan.split("-").map(Number);
            const now = new Date();
            const limit =
              now.getFullYear() === y && now.getMonth() + 1 === m
                ? now.getDate()
                : undefined;
            const effCount = getEffectiveDays(
              y,
              m - 1,
              holidays,
              settings,
              limit,
            ).length;
            const perc =
              effCount > 0
                ? Math.round(((hadir + izin + sakit) / effCount) * 100)
                : 0;
            return {
              "No.": idx + 1,
              "Nama Siswa": s.nama,
              Kelas: s.kelas,
              Hadir: hadir,
              Izin: izin,
              Sakit: sakit,
              Alfa: alfa,
              "Terlambat (Menit)": totalLambat,
              "Jam Terlambat": formatMinutes(totalLambat),
              "% Kehadiran": `${perc}%`,
            };
          })
      : teachers
          .filter((t) => t.nip !== "ADMIN001" && t.role !== "Admin")
          .filter((t) => {
            if (!rekapFilter.kelas) return true;
            const scheds = teachingSchedules.filter((ts) => ts.nip === t.nip);
            return scheds.some((sc) => sc.kelas === rekapFilter.kelas);
          })
          .map((t, idx) => {
            const schedules = teachingSchedules.filter(
              (ts) =>
                ts.nip === t.nip &&
                (rekapFilter.kelas ? ts.kelas === rekapFilter.kelas : true),
            );
            const { totalTarget, hadir, sakit, izin, alfa, totalLambat, perc } =
              calculateTeacherMonthlyStats(
                t.nip,
                rekapFilter.bulan,
                rekapFilter.kelas,
                teachingSchedules,
                holidays,
                settings,
                teacherAttendance,
              );
            const mapelTaught =
              Array.from(
                new Set(schedules.map((sc) => sc.mapel).filter(Boolean)),
              ).join(", ") || "-";
            return {
              "No.": idx + 1,
              "Nama Guru": t.nama,
              Mapel: mapelTaught,
              Kelas:
                rekapFilter.kelas ||
                schedules.map((s) => s.kelas).join(", ") ||
                "-",
              "Target Sesi": totalTarget,
              Hadir: hadir,
              Sakit: sakit,
              Izin: izin,
              Alfa: alfa,
              "Terlambat (Menit)": totalLambat,
              "Jam Terlambat": formatMinutes(totalLambat),
              "% Kinerja": `${perc}% (Hadir: ${hadir}, Sakit: ${sakit}, Izin: ${izin}, Alpa: ${alfa})`,
            };
          });

    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Rekap Kehadiran");
    XLSX.writeFile(wb, `Rekap-${rekapFilter.type}-${rekapFilter.bulan}.xlsx`);
  };

  const exportPDF = () => {
    const doc = new jsPDF() as any;
    const isSiswa = rekapFilter.type === "Siswa";
    const data = isSiswa
      ? students
          .filter((s) =>
            rekapFilter.kelas ? s.kelas === rekapFilter.kelas : true,
          )
          .map((s, idx) => {
            const m = attendance.filter(
              (a) =>
                a.nisn === s.nisn && a.tanggal.startsWith(rekapFilter.bulan),
            );
            const h = m.filter((a) => a.status === "Hadir").length;
            const i = m.filter((a) => a.status === "Izin").length;
            const sk = m.filter((a) => a.status === "Sakit").length;
            const f = m.filter((a) => a.status === "Alfa").length;
            const lbt = m.reduce((sum, a) => sum + (a.terlambat || 0), 0);
            const [y, mm] = rekapFilter.bulan.split("-").map(Number);
            const now = new Date();
            const limit =
              now.getFullYear() === y && now.getMonth() + 1 === mm
                ? now.getDate()
                : undefined;
            const effCount = getEffectiveDays(
              y,
              mm - 1,
              holidays,
              settings,
              limit,
            ).length;
            const p =
              effCount > 0 ? Math.round(((h + i + sk) / effCount) * 100) : 0;
            return [idx + 1, s.nama, s.kelas, h, i, sk, f, lbt, `${p}%`];
          })
      : teachers
          .filter((t) => t.nip !== "ADMIN001" && t.role !== "Admin")
          .filter((t) => {
            if (!rekapFilter.kelas) return true;
            const scheds = teachingSchedules.filter((ts) => ts.nip === t.nip);
            return scheds.some((sc) => sc.kelas === rekapFilter.kelas);
          })
          .map((t, idx) => {
            const schs = teachingSchedules.filter(
              (ts) =>
                ts.nip === t.nip &&
                (rekapFilter.kelas ? ts.kelas === rekapFilter.kelas : true),
            );
            const { totalTarget, hadir, sakit, izin, alfa, perc } =
              calculateTeacherMonthlyStats(
                t.nip,
                rekapFilter.bulan,
                rekapFilter.kelas,
                teachingSchedules,
                holidays,
                settings,
                teacherAttendance,
              );
            const mapelTaught =
              Array.from(
                new Set(schs.map((sc) => sc.mapel).filter(Boolean)),
              ).join(", ") || "-";
            return [
              idx + 1,
              t.nama,
              mapelTaught,
              rekapFilter.kelas || schs.map((s) => s.kelas).join(", ") || "-",
              totalTarget,
              hadir,
              sakit,
              izin,
              alfa,
              `${perc}%`,
            ];
          });

    const getIndonesianMonthYear = (bulanStr: string) => {
      if (!bulanStr) return "";
      const parts = bulanStr.split("-");
      if (parts.length === 2) {
        const year = parts[0];
        const monthNum = parseInt(parts[1], 10);
        const indonesianMonths = [
          "Januari",
          "Februari",
          "Maret",
          "April",
          "Mei",
          "Juni",
          "Juli",
          "Agustus",
          "September",
          "Oktober",
          "November",
          "Desember",
        ];
        const monthName = indonesianMonths[monthNum - 1] || parts[1];
        return `${monthName}/${year}`;
      }
      return bulanStr;
    };

    const classString = rekapFilter.kelas ? ` Kelas ${rekapFilter.kelas}` : "";
    const formattedBulan = getIndonesianMonthYear(rekapFilter.bulan);

    doc.setFontSize(14);
    doc.text(
      `Laporan Rekap Kehadiran ${rekapFilter.type}${classString}`,
      14,
      15,
    );
    doc.setFontSize(10);
    doc.text(`Bulan: ${formattedBulan}`, 14, 23);

    const columns = isSiswa
      ? ["No.", "Nama", "Kelas", "H", "I", "S", "A", "Lbt(m)", "%"]
      : [
          "No.",
          "Nama Guru",
          "Mapel",
          "Kelas",
          "Target",
          "H",
          "S",
          "I",
          "A",
          "%",
        ];

    autoTable(doc, {
      startY: 28,
      head: [columns],
      body: data,
    });
    doc.save(`Rekap-${rekapFilter.type}-${rekapFilter.bulan}.pdf`);
  };

  const toggleLoader = (val: boolean) => setLoading(val);
  const triggerCurrentPanelRefresh = async () => {
    setRefreshing(true);
    refreshMasterData();
    // Reset all filters and pagination to ensure user sees latest data
    setSearchTermGuru("");
    setSearchTermSiswa("");
    setSearchTermKelas("");
    setSearchTermJadwal("");
    setFilterClassAbsensi("");
    setFilterNamaAbsensi("");
    setFilterAdminSiswaClass("");
    setStudentProfileClassFilter("");
    setAnalysisClass("");

    // Reset pagination
    setPagination({
      guru: 0,
      siswa: 0,
      kelas: 0,
      jadwal: 0,
      absensiSiswa: 0,
      absensiGuru: 0,
      rekapMapel: 0,
    });

    // Real-time data is synced via onSnapshot, but we can simulate a fresh state
    setLoading(true);
    setTimeout(() => {
      setRefreshing(false);
      setLoading(false);
      triggerSuccess(
        "DIPERBARUI",
        "Sinkronisasi data terbaru berhasil dilakukan. Semua filter telah direset.",
      );
    }, 1000);
  };

  const handleLogin = async () => {
    const u = loginRole === "Siswa" ? loginNisn : loginUser;
    const p = loginRole === "Siswa" ? "" : loginPass;

    if (!u || (loginRole !== "Siswa" && !p)) {
      const msg =
        loginRole === "Siswa"
          ? "Nomor Induk Siswa Nasional (NISN) wajib diisi."
          : "ID Pengguna dan Kata Sandi wajib diisi.";
      setLoginError(msg);
      alert(msg);
      return;
    }

    setLoginError("");
    toggleLoader(true);
    try {
      const res = await firestoreService.checkLogin(u, p, loginRole);
      toggleLoader(false);
      if (res.success) {
        setSession(res);
        if (res.role === "Siswa") setActivePanel("siswa-personal");
      } else {
        const errorMsg =
          res.message ||
          (loginRole === "Siswa"
            ? "NISN salah atau tidak terdaftar."
            : "Username atau Password salah.");
        setLoginError(errorMsg);
        alert(errorMsg);
      }
    } catch (e) {
      toggleLoader(false);
      const errMsg =
        "Gagal terhubung ke database. Silakan coba beberapa saat lagi.";
      setLoginError(errMsg);
      alert(errMsg);
    }
  };

  const defaultLogo = "/logo.png";
  const fallbackLogo =
    "https://www.kemenag.go.id/assets/images/logo-kemenag.png";

  const appLogo = useMemo(() => {
    return appConfig.logoUrl || defaultLogo;
  }, [appConfig.logoUrl]);

  const dayOrder: any = {
    Senin: 1,
    Selasa: 2,
    Rabu: 3,
    Kamis: 4,
    Jumat: 5,
    Sabtu: 6,
    Minggu: 7,
  };

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
                Sistem mendeteksi bahwa batas kuota gratis harian (
                <strong>Firestore Free Tier Read units</strong>) untuk proyek
                database Firebase ini telah terlampaui.
              </p>
              <div className="flex gap-2.5 items-start">
                <div className="w-2 h-2 rounded-full bg-red-500 mt-1.5 shrink-0" />
                <p>
                  Batas harian Spark Plan adalah{" "}
                  <strong>50.000 pembacaan</strong> per hari. Layanan akan
                  dipulihkan secara otomatis setelah kuota di-reset kembali
                  keesokan harinya oleh Google Cloud.
                </p>
              </div>
              <div className="flex gap-2.5 items-start">
                <div className="w-2 h-2 rounded-full bg-zinc-900 mt-1.5 shrink-0" />
                <p>
                  <strong>Untuk Pemilik Proyek:</strong> Anda dapat mengaktifkan
                  billing (upgrade ke Blaze/Pay-as-you-go plan) pada konsol
                  Firebase untuk menghindari batas harian ini sama sekali.
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
                <ExternalLink
                  size={16}
                  className="text-amber-400 group-hover:rotate-12 transition-transform"
                />
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
              *Proyek ID: gen-lang-client-0541836694 • DB:
              ai-studio-8c6b2059-8af6-4d53-bb89-22a78bab9d06
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
            <RefreshCcw
              className="animate-spin text-green-700 mb-2"
              size={32}
            />
            <p className="font-bold text-green-905 text-xs uppercase tracking-widest">
              {loaderText || "Sinkronisasi..."}
            </p>
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
                  target.style.opacity = "0";
                  const parent = target.parentElement;
                  if (parent) {
                    const icon = parent.querySelector(".fallback-icon");
                    if (icon) (icon as HTMLElement).style.opacity = "1";
                  }
                } else {
                  target.src = defaultLogo;
                }
              }}
            />
            <School
              size={48}
              className="text-green-700 absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 opacity-0 fallback-icon transition-opacity"
            />
          </div>
          <h1 className="text-4xl font-black tracking-tighter">SIGAP</h1>
          <p className="text-green-100 text-[10px] font-bold uppercase tracking-widest mt-1 opacity-80">
            Sistem Informasi Guru dan Absensi Pelajar
          </p>
          <div className="mt-3 inline-block">
            <p className="text-white font-black text-xs bg-green-900/40 py-1.5 rounded-full border border-white/20 px-5 shadow-lg backdrop-blur-sm uppercase tracking-wider">
              MTsN 2 BOMBANA
            </p>
          </div>

          <div className="mt-4 flex flex-col items-center justify-center gap-1">
            <div className="flex items-center gap-2">
              <div
                className={`w-2 h-2 rounded-full animate-pulse ${firebaseConnected ? "bg-green-400 shadow-[0_0_8px_#4ade80]" : "bg-red-500"}`}
              />
              <span className="text-[9px] font-bold text-white/60 uppercase tracking-widest">
                Database:{" "}
                {firebaseConnected
                  ? "Terhubung (Online)"
                  : "Terputus (Offline)"}
              </span>
            </div>
            {quotaExceeded && (
              <div className="bg-red-500/20 px-3 py-1.5 rounded-lg border border-red-500/30 mt-2">
                <p className="text-[9px] font-black text-red-200 uppercase tracking-widest animate-pulse flex items-center gap-2">
                  <AlertCircle size={10} /> Quota Firebase Habis (Limit
                  Tercapai)
                </p>
              </div>
            )}
          </div>
        </div>

        <div className="p-8">
          {(() => {
            const appRole = import.meta.env.VITE_APP_ROLE || "staff";
            const roles = (["Siswa", "Guru", "Admin"] as const).filter((r) => {
              if (appRole === "student") return r === "Siswa";
              if (appRole === "teacher") return r === "Siswa" || r === "Guru";
              return true;
            });

            if (roles.length <= 1) return null;

            return (
              <div className="flex bg-gray-100 p-1 rounded-xl mb-6">
                {roles.map((role) => (
                  <button
                    key={role}
                    disabled={loading}
                    onClick={() => setLoginRole(role)}
                    className={`flex-1 py-2 text-sm font-bold rounded-lg transition-all ${loginRole === role ? "bg-white text-green-800 shadow-sm" : "text-gray-500"}`}
                  >
                    {role}
                  </button>
                ))}
              </div>
            );
          })()}

          <div className="space-y-4">
            <AnimatePresence>
              {anonymousAuthError && (
                <motion.div
                  initial={{ opacity: 0, y: -10 }}
                  animate={{ opacity: 1, y: 0 }}
                  className="bg-amber-50 border border-amber-200 rounded-2xl p-4 flex flex-col gap-2.5 text-amber-900 shadow-sm"
                >
                  <div className="flex items-start gap-3 text-left">
                    <AlertTriangle
                      size={20}
                      className="text-amber-600 shrink-0 mt-0.5 animate-bounce"
                    />
                    <div className="text-xs font-semibold leading-relaxed">
                      <span className="font-extrabold uppercase tracking-wider block text-[10px] text-amber-800 mb-0.5">
                        Konfigurasi Firebase Diperlukan
                      </span>
                      Sistem mendeteksi bahwa <strong>Login Anonim (Anonymous Auth)</strong> belum diaktifkan di konsol Firebase Anda. Langkah masuk menggunakan database yang aman memerlukan otentikasi ini.
                    </div>
                  </div>
                  <div className="text-[11px] text-left bg-white/60 p-3 rounded-lg border border-amber-100 font-medium space-y-1.5 text-zinc-700">
                    <p className="font-bold text-amber-950 uppercase tracking-wide text-[9px]">Cara Mengaktifkan di Konsol Firebase Anda:</p>
                    <ol className="list-decimal list-inside space-y-1">
                      <li>Masuk ke <strong>Firebase Console</strong> proyek Anda.</li>
                      <li>Di menu samping, pilih <strong>Build &gt; Authentication</strong>.</li>
                      <li>Buka tab <strong>Sign-in method</strong>.</li>
                      <li>Klik <strong>Add new provider</strong>, lalu pilih <strong>Anonymous (Anonim)</strong>.</li>
                      <li>Aktifkan tombol geser (Enable) lalu klik <strong>Save</strong>.</li>
                    </ol>
                  </div>
                </motion.div>
              )}
            </AnimatePresence>

            <AnimatePresence>
              {loginError && (
                <motion.div
                  initial={{ opacity: 0, y: -10 }}
                  animate={{ opacity: 1, y: 0 }}
                  exit={{ opacity: 0, y: -10 }}
                  className="bg-red-50 border border-red-200 rounded-xl p-4 flex items-start gap-3 text-red-800"
                >
                  <AlertCircle
                    size={18}
                    className="text-red-600 shrink-0 mt-0.5"
                  />
                  <div className="text-xs font-semibold leading-relaxed">
                    <span className="font-extrabold uppercase tracking-wider block text-[10px] text-red-700 mb-0.5">
                      Kesalahan Masuk
                    </span>
                    {loginError}
                  </div>
                </motion.div>
              )}
            </AnimatePresence>

            {loginRole === "Siswa" ? (
              <div>
                <label className="block text-xs font-bold text-gray-400 uppercase mb-1 ml-1">
                  Nomor Induk Siswa Nasional (NISN)
                </label>
                <input
                  type="number"
                  disabled={loading}
                  value={loginNisn}
                  onChange={(e) => {
                    setLoginNisn(e.target.value);
                    setLoginError("");
                  }}
                  className="w-full bg-gray-50 border-0 rounded-xl py-3 px-4 focus:ring-2 focus:ring-green-500 transition-all font-medium disabled:opacity-50"
                  placeholder="Masukkan NISN"
                />
              </div>
            ) : (
              <>
                <div>
                  <label className="block text-xs font-bold text-gray-400 uppercase mb-1 ml-1">
                    ID Pengguna / NIP
                  </label>
                  <input
                    type="text"
                    disabled={loading}
                    value={loginUser}
                    onChange={(e) => {
                      setLoginUser(e.target.value);
                      setLoginError("");
                    }}
                    className="w-full bg-gray-50 border-0 rounded-xl py-3 px-4 focus:ring-2 focus:ring-green-500 transition-all font-medium disabled:opacity-50"
                    placeholder="Username"
                  />
                </div>
                <div>
                  <label className="block text-xs font-bold text-gray-400 uppercase mb-1 ml-1">
                    Kata Sandi
                  </label>
                  <div className="relative">
                    <input
                      type={showPass ? "text" : "password"}
                      disabled={loading}
                      value={loginPass}
                      onChange={(e) => {
                        setLoginPass(e.target.value);
                        setLoginError("");
                      }}
                      className="w-full bg-gray-50 border-0 rounded-xl py-3 px-4 focus:ring-2 focus:ring-green-500 transition-all font-medium disabled:opacity-50"
                      placeholder="••••••••"
                    />
                    <button
                      onClick={() => setShowPass(!showPass)}
                      className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-400 hover:text-green-800 transition-colors"
                    >
                      {showPass ? <EyeOff size={18} /> : <Eye size={18} />}
                    </button>
                  </div>
                </div>
              </>
            )}

            <button
              onClick={handleLogin}
              disabled={loading}
              className={`w-full bg-green-800 text-white rounded-xl py-4 font-bold shadow-lg hover:bg-green-700 active:scale-95 transition-all mt-4 flex items-center justify-center gap-2 ${loading ? "opacity-70 cursor-wait" : ""}`}
            >
              {loading ? (
                <>
                  <RefreshCcw className="animate-spin" size={20} />
                  Memproses...
                </>
              ) : (
                "Masuk Aplikasi"
              )}
            </button>
          </div>
        </div>
      </motion.div>
    </div>
  );

  const [scanInput, setScanInput] = useState("");

  const handleManualScan = async () => {
    if (!scanInput) return;
    toggleLoader(true);
    try {
      let res;
      if (navigator.onLine) {
        try {
          if (scanType === "Siswa") {
            res = await firestoreService.processScan(scanInput);
          } else {
            res = await firestoreService.processClassScan(
              session?.uid || "",
              scanInput,
            );
          }
        } catch (e) {
          console.warn("Firestore scan error, falling back to offline mode:", e);
          res = processOfflineScan(scanInput, scanType);
        }
      } else {
        res = processOfflineScan(scanInput, scanType);
      }

      if (res.success) {
        const audio = new Audio(
          "https://assets.mixkit.co/active_storage/sfx/2568/2568-preview.mp3",
        );
        audio.play().catch((e) => console.log("Audio play blocked"));
        if (navigator.vibrate) navigator.vibrate(200);
      }

      setScanResult({
        success: res.success,
        message: res.message,
        status: res.status,
      });
      setScanInput("");
    } catch (e) {
      setScanResult({ success: false, message: "Gagal memproses pindaian." });
    } finally {
      toggleLoader(false);
    }
  };

  const handleLogout = () => {
    firestoreService.logout();
    setSession(null);
    setLoginNisn("");
    setLoginUser("");
    setLoginPass("");
    setActivePanel("home");
  };

  const handleImportExcel = (type: "Siswa" | "Guru") => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = ".xlsx, .xls";
    input.onchange = (e: any) => {
      const file = e.target.files[0];
      const reader = new FileReader();
      reader.onload = async (evt: any) => {
        const bstr = evt.target.result;
        const wb = XLSX.read(bstr, { type: "binary" });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws);

        toggleLoader(true);
        try {
          const promises = (data as any[]).map((item) => {
            if (type === "Siswa") return firestoreService.saveSiswa(item);
            return firestoreService.saveGuru({
              ...item,
              role: (item.role || "Guru") as any,
            });
          });
          await Promise.all(promises);
          triggerSuccess(
            "BERHASIL",
            `Sukses mengimpor ${data.length} data ${type}. Data akan muncul otomatis di tabel.`,
          );
        } catch (err) {
          triggerError(
            "GAGAL",
            "Gagal mengimpor data. Pastikan format template sesuai dan data tidak duplikat.",
          );
        } finally {
          toggleLoader(false);
        }
      };
      reader.readAsBinaryString(file);
    };
    input.click();
  };

  const downloadTemplate = (type: "Siswa" | "Guru") => {
    const headers =
      type === "Siswa"
        ? [["nisn", "nama", "tempat", "tgl", "kelas", "ayah", "ibu", "hp"]]
        : [["nip", "nama", "kelas", "user", "pass"]];
    const ws = XLSX.utils.aoa_to_sheet(headers);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Template");
    XLSX.writeFile(wb, `Template_Import_${type}.xlsx`);
  };

  const getSiswaPercentage = useMemo(() => {
    if (session?.role !== "Siswa") return 0;
    const today = new Date();
    const month = today.getMonth() + 1;
    const year = today.getFullYear();
    const currentMonthAbsen = attendance.filter((a) => {
      const [aYear, aMonth] = a.tanggal.split("-").map(Number);
      return a.nisn === session?.uid && aMonth === month && aYear === year;
    });

    if (currentMonthAbsen.length === 0 && effectiveDaysThisMonth.length === 0)
      return 0;
    const hadir = currentMonthAbsen.filter(
      (a) =>
        a.status === "Hadir" || a.status === "Izin" || a.status === "Sakit",
    ).length;
    const effCount = effectiveDaysThisMonth.length;
    return effCount > 0 ? Math.round((hadir / effCount) * 100) : 0;
  }, [attendance, session, effectiveDaysThisMonth]);

  const menuItems = useMemo(() => {
    if (!session) return [];
    if (session?.role === "Admin") {
      return [
        { id: "home", label: "Dashboard Utama", icon: LayoutGrid },
        { id: "scanner", label: "Scanner Presensi", icon: QrCode },
        { id: "guru", label: "Data Guru", icon: UserPlus },
        { id: "siswa", label: "Data Siswa", icon: GraduationCap },
        { id: "kelas", label: "Data Kelas", icon: TableIcon },
        { id: "jadwal-mengajar", label: "Jam Mengajar", icon: Calendar },
        { id: "absensi-umum", label: "Absensi Siswa", icon: ClipboardList },
        { id: "absensi-guru", label: "Absensi Guru", icon: User },
        { id: "rekap", label: "Laporan Rekap", icon: FileBarChart },
        { id: "rekap-mapel", label: "Rekap Mapel", icon: ClipboardList },
        { id: "analisis", label: "Analisis Kehadiran", icon: BarChart3 },
        { id: "pengaturan", label: "Pengaturan", icon: Settings },
      ];
    } else if (session?.role === "Guru") {
      const teacher = teachers.find((t) => t.nip === session?.uid);
      const jabatan = teacher?.jabatan || "Guru";

      const items = [
        { id: "home", label: "Dashboard Utama", icon: LayoutGrid },
        { id: "scanner", label: "Scanner Presensi", icon: QrCode },
      ];

      items.push({ id: "profil", label: "Profil Saya", icon: User });

      if (jabatan === "Kamad") {
        items.push({
          id: "absensi-umum",
          label: "Absensi Siswa",
          icon: ClipboardList,
        });
        items.push({ id: "absensi-guru", label: "Absensi Guru", icon: User });
        items.push({
          id: "input-kehadiran-guru",
          label: "Input Kehadiran Guru",
          icon: CalendarRange,
        });
        items.push({
          id: "capaian-guru",
          label: "Laporan Capaian Guru",
          icon: Award,
        });
        items.push({
          id: "analisis",
          label: "Analisis Kehadiran",
          icon: BarChart3,
        });
      } else if (jabatan === "Wakamad") {
        items.push({
          id: "absensi-umum",
          label: "Absensi Siswa",
          icon: ClipboardList,
        });
        items.push({ id: "absensi-guru", label: "Absensi Guru", icon: User });
        items.push({ id: "roster", label: "Roster Mengajar", icon: TableIcon });
        items.push({
          id: "profil-siswa",
          label: "Profil Siswa",
          icon: GraduationCap,
        });
        items.push({ id: "rekap", label: "Laporan Rekap", icon: FileBarChart });
        items.push({
          id: "rekap-mapel",
          label: "Rekap Mapel",
          icon: ClipboardList,
        });
        items.push({
          id: "analisis",
          label: "Analisis Kehadiran",
          icon: BarChart3,
        });
      } else if (jabatan === "Guru Wali Kelas" || session?.isWali) {
        items.push({
          id: "profil-siswa",
          label: "Profil Siswa",
          icon: GraduationCap,
        });
        items.push({
          id: "absensi-wali",
          label: "Riwayat Kehadiran",
          icon: ClipboardList,
        });
        items.push({ id: "rekap", label: "Laporan Rekap", icon: FileBarChart });
        items.push({
          id: "rekap-mapel",
          label: "Rekap Mapel",
          icon: ClipboardList,
        });
        items.push({
          id: "analisis",
          label: "Analisis Kehadiran",
          icon: BarChart3,
        });
      } else {
        items.push({
          id: "rekap-mapel",
          label: "Rekap Mapel",
          icon: ClipboardList,
        });
      }

      items.push({ id: "kbc", label: "KBC", icon: Heart });

      return items;
    } else {
      return [
        { id: "siswa-personal", label: "Dashboard Siswa", icon: User },
        { id: "siswa-data", label: "Data Siswa", icon: GraduationCap },
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
                target.style.opacity = "0";
                const parent = target.parentElement;
                if (parent) {
                  const icon = parent.querySelector(".fallback-icon");
                  if (icon) (icon as HTMLElement).style.opacity = "1";
                }
              } else {
                target.src = defaultLogo;
              }
            }}
          />
          <School
            size={28}
            className="text-green-700 absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 opacity-0 fallback-icon transition-opacity"
          />
        </div>
        <h2 className="font-black text-xs text-center text-white uppercase tracking-widest">
          SIGAP MTsN 2
        </h2>
        <span className="bg-white/10 text-white text-[9px] px-2 py-0.5 rounded-full font-bold mt-2 uppercase border border-white/10">
          {session?.jabatan || session?.role}{" "}
          {session?.isWali && `(Wali ${session.kelas})`}
        </span>
      </div>

      <nav className="flex-grow py-4 overflow-y-auto">
        {menuItems.map((item) => (
          <button
            key={item.id}
            onClick={() => setActivePanel(item.id)}
            className={`w-full flex items-center px-6 py-4 transition-all text-sm font-medium border-l-4 ${activePanel === item.id ? "bg-green-900/20 text-white border-green-600" : "text-zinc-400 border-transparent hover:text-white hover:bg-zinc-900"}`}
          >
            <item.icon size={18} className="mr-3" />
            {item.label}
          </button>
        ))}
      </nav>

      <div className="p-4 border-t border-zinc-800 space-y-3">
        <div className="flex items-center justify-center gap-2 p-2 bg-zinc-900 rounded-lg border border-zinc-800">
          <div
            className={`w-2 h-2 rounded-full ${firebaseConnected ? "bg-green-500 shadow-[0_0_8px_rgba(34,197,94,0.6)]" : "bg-red-500"}`}
          />
          <span className="text-[10px] font-bold text-zinc-500 uppercase tracking-widest">
            Firebase: {firebaseConnected ? "Connected" : "Disconnected"}
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
      refreshMasterData();
      setShowSiswaModal(false);
      setEditingSiswa(null);
      setSearchTermSiswa(""); // Reset search to show newcomer
      triggerSuccess(
        "BERHASIL",
        "Data siswa telah berhasil disimpan. Data akan muncul otomatis di daftar.",
      );
    } catch (e) {
      alert("Gagal menyimpan data.");
    } finally {
      toggleLoader(false);
    }
  };

  const handleDeleteSiswa = (nisn: string, nama: string) => {
    setConfirmModal({
      show: true,
      title: "Hapus data siswa?",
      message:
        "Data siswa dan riwayat absensinya akan dihapus permanen dari sistem.",
      entityName: nama,
      onConfirm: async () => {
        toggleLoader(true);
        try {
          await firestoreService.hapusSiswa(nisn);
          refreshMasterData();
        } catch (e) {
          alert("Gagal menghapus data.");
        } finally {
          toggleLoader(false);
        }
      },
    });
  };

  const handleSaveGuru = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingGuru) return;
    toggleLoader(true);
    // Ensure kelas is only for Wali Kelas, and handle empty user/pass
    const cleanedGuru = {
      ...editingGuru,
      jabatan: editingGuru.jabatan || "Guru",
      role: editingGuru.role || "Guru",
      user: editingGuru.user?.trim() || `guru${editingGuru.nip}`.toLowerCase(),
      pass: editingGuru.pass?.trim() || "123456",
      kelas: editingGuru.jabatan === "Guru Wali Kelas" ? editingGuru.kelas : "",
    };
    try {
      await firestoreService.saveGuru(cleanedGuru);
      refreshMasterData();
      setShowGuruModal(false);
      setEditingGuru(null);
      setSearchTermGuru(""); // Reset search to show latest data
      triggerSuccess(
        "BERHASIL",
        "Data guru dan akun akses telah diperbarui. Perubahan akan muncul otomatis.",
      );
    } catch (e) {
      alert("Gagal menyimpan data.");
    } finally {
      toggleLoader(false);
    }
  };

  const handleDeleteGuru = (nip: string, nama: string) => {
    setConfirmModal({
      show: true,
      title: "Hapus guru?",
      message: "Data guru dan akses loginnya akan dihapus permanen.",
      entityName: nama,
      onConfirm: async () => {
        toggleLoader(true);
        try {
          await firestoreService.hapusGuru(nip);
          refreshMasterData();
        } catch (e) {
          alert("Gagal menghapus data.");
        } finally {
          toggleLoader(false);
        }
      },
    });
  };

  const handleSaveKelas = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingKelas) return;
    toggleLoader(true);
    try {
      await firestoreService.saveKelas(editingKelas);
      refreshMasterData();
      setShowKelasModal(false);
      setEditingKelas(null);
      setSearchTermKelas(""); // Reset search
      triggerSuccess(
        "BERHASIL",
        `Kelas ${editingKelas.nama} berhasil disimpan dan sinkron.`,
      );
    } catch (e) {
      alert("Gagal menyimpan data.");
    } finally {
      toggleLoader(false);
    }
  };

  const handleDeleteKelas = (nama: string) => {
    setConfirmModal({
      show: true,
      title: "Hapus kelas?",
      message:
        "Data kelas akan dihapus. Siswa yang terdaftar di kelas ini mungkin perlu dipindahkan.",
      entityName: nama,
      onConfirm: async () => {
        toggleLoader(true);
        try {
          await firestoreService.hapusKelas(nama);
          refreshMasterData();
        } catch (e) {
          alert("Gagal menghapus data.");
        } finally {
          toggleLoader(false);
        }
      },
    });
  };

  const renderNaikKelasModal = () => (
    <AnimatePresence>
      {showNaikKelasModal && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center p-4">
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="absolute inset-0 bg-black/60 backdrop-blur-sm"
            onClick={() => setShowNaikKelasModal(false)}
          />
          <motion.div
            initial={{ scale: 0.9, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            exit={{ scale: 0.9, opacity: 0 }}
            className="relative bg-white w-full max-w-2xl rounded-3xl shadow-2xl overflow-hidden flex flex-col max-h-[90vh]"
          >
            <div className="bg-gradient-to-r from-orange-600 to-orange-500 p-6 text-white text-center flex flex-col items-center">
              <ArrowUpCircle size={36} className="mb-2 animate-bounce" />
              <h2 className="text-xl font-bold uppercase tracking-wide">
                Kenaikan Kelas Massal
              </h2>
              <p className="text-orange-100 text-xs mt-1">
                Ubah kelas para siswa sekaligus dengan cepat
              </p>
            </div>

            <form
              onSubmit={handleProsesNaikKelas}
              className="p-6 overflow-y-auto space-y-6 custom-scrollbar flex-grow"
            >
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">
                    1. Pilih Kelas Asal
                  </label>
                  <select
                    required
                    value={naikKelasOrigin}
                    onChange={(e) => {
                      setNaikKelasOrigin(e.target.value);
                      setNaikKelasSearchKeyword("");
                    }}
                    className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold text-sm mt-1 focus:ring-2 focus:ring-orange-500"
                  >
                    <option value="">-- Pilih Kelas Asal --</option>
                    {naikKelasOriginOptions.map((clsName) => (
                      <option key={clsName} value={clsName}>
                        {clsName}
                      </option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">
                    2. Cari / Filter Nama Siswa
                  </label>
                  <div className="relative mt-1">
                    <Search
                      size={16}
                      className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400"
                    />
                    <input
                      type="text"
                      placeholder="Cari nama atau NISN..."
                      disabled={!naikKelasOrigin}
                      value={naikKelasSearchKeyword}
                      onChange={(e) =>
                        setNaikKelasSearchKeyword(e.target.value)
                      }
                      className="w-full bg-zinc-50 border-0 rounded-xl py-3 pl-10 pr-4 text-sm font-bold focus:ring-2 focus:ring-orange-500 disabled:opacity-50"
                    />
                  </div>
                </div>
              </div>

              {naikKelasOrigin && (
                <div className="space-y-3">
                  <div className="flex justify-between items-center bg-orange-50 px-4 py-3 rounded-xl border border-orange-100">
                    <p className="text-xs font-bold text-orange-950 uppercase tracking-wider">
                      Daftar Siswa ({naikKelasFilteredStudents.length}{" "}
                      Terdeteksi)
                    </p>
                    {naikKelasFilteredStudents.length > 0 && (
                      <button
                        type="button"
                        onClick={handleToggleSelectAllNaikKelas}
                        className="text-[10px] font-black uppercase text-orange-700 bg-white hover:bg-orange-100 border border-orange-200 px-3 py-1.5 rounded-lg transition-all"
                      >
                        {naikKelasFilteredStudents.every((s) =>
                          naikKelasSelectedStudents.includes(s.nisn),
                        )
                          ? "Batal Pilih Semua"
                          : "Pilih Semua"}
                      </button>
                    )}
                  </div>

                  <div className="max-h-60 overflow-y-auto border border-gray-100 rounded-2xl divide-y divide-gray-50 bg-white custom-scrollbar w-full">
                    {naikKelasFilteredStudents.length > 0 ? (
                      naikKelasFilteredStudents.map((s) => {
                        const isSelected = naikKelasSelectedStudents.includes(
                          s.nisn,
                        );
                        return (
                          <div
                            key={s.nisn}
                            onClick={() =>
                              handleToggleSelectStudentNaikKelas(s.nisn)
                            }
                            className={`flex items-center gap-3 p-3 cursor-pointer hover:bg-orange-50/50 transition-colors ${isSelected ? "bg-orange-50/20" : ""}`}
                          >
                            <input
                              type="checkbox"
                              checked={isSelected}
                              onChange={() => {}} // Handled by parent click
                              className="rounded text-orange-600 focus:ring-orange-500 w-4 h-4"
                            />
                            {s.foto ? (
                              <img
                                src={s.foto}
                                className="w-8 h-8 rounded-lg object-cover"
                                alt=""
                              />
                            ) : (
                              <div className="w-8 h-8 rounded-lg bg-gray-100 flex items-center justify-center text-gray-400">
                                <User size={14} />
                              </div>
                            )}
                            <div className="flex-grow">
                              <p className="text-xs font-extrabold text-zinc-900">
                                {s.nama}
                              </p>
                              <p className="text-[10px] font-bold text-gray-400 font-mono">
                                NISN: {s.nisn}
                              </p>
                            </div>
                            <span className="text-[10px] bg-zinc-100 text-zinc-600 font-bold px-2 py-0.5 rounded-full">
                              {s.kelas}
                            </span>
                          </div>
                        );
                      })
                    ) : (
                      <div className="p-8 text-center text-gray-400 text-xs italic">
                        {naikKelasSearchKeyword
                          ? "Tidak ada siswa yang cocok dengan pencarian."
                          : "Tidak ada siswa yang terdaftar di kelas ini."}
                      </div>
                    )}
                  </div>

                  <div className="text-right">
                    <span className="text-[10px] bg-orange-100 text-orange-850 font-black px-3 py-1.5 rounded-full uppercase tracking-wider">
                      {naikKelasSelectedStudents.length} Siswa Terpilih
                    </span>
                  </div>
                </div>
              )}

              <div className="bg-zinc-50 p-4 rounded-2xl space-y-3">
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                  <div className="flex items-center gap-2 text-zinc-700">
                    <ArrowUpCircle size={20} className="text-orange-600" />
                    <div>
                      <h4 className="text-xs font-black uppercase tracking-wider text-orange-950">
                        3. Tentukan Kelas Tujuan
                      </h4>
                      <p className="text-[10px] text-zinc-400">
                        Siswa terpilih akan otomatis dipindahkan ke kelas ini
                      </p>
                    </div>
                  </div>

                  <div className="w-full md:w-64">
                    <select
                      required
                      disabled={
                        !naikKelasOrigin ||
                        naikKelasSelectedStudents.length === 0
                      }
                      value={naikKelasTarget}
                      onChange={(e) => setNaikKelasTarget(e.target.value)}
                      className="w-full bg-white border border-gray-200 rounded-xl px-4 py-3 font-bold text-sm focus:ring-2 focus:ring-orange-500 disabled:opacity-50"
                    >
                      <option value="">-- Pilih Kelas Tujuan --</option>
                      {naikKelasTargetOptions.map((clsName) => (
                        <option
                          key={clsName}
                          disabled={clsName === naikKelasOrigin}
                          value={clsName}
                        >
                          {clsName}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>
              </div>

              <div className="pt-2 flex gap-3">
                <button
                  type="button"
                  onClick={() => setShowNaikKelasModal(false)}
                  className="flex-1 py-4 font-bold text-zinc-400"
                >
                  Batal
                </button>
                <button
                  type="submit"
                  disabled={
                    !naikKelasOrigin ||
                    !naikKelasTarget ||
                    naikKelasSelectedStudents.length === 0
                  }
                  className="flex-1 bg-gradient-to-r from-orange-600 to-orange-500 text-white rounded-2xl py-4 font-bold hover:from-orange-500 hover:to-orange-400 shadow-xl disabled:opacity-40 disabled:cursor-not-allowed"
                >
                  Proses Naik Kelas
                </button>
              </div>
            </form>
          </motion.div>
        </div>
      )}
    </AnimatePresence>
  );

  const renderSiswaModal = () => (
    <AnimatePresence>
      {showSiswaModal && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center p-4">
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="absolute inset-0 bg-black/60 backdrop-blur-sm"
            onClick={() => setShowSiswaModal(false)}
          />
          <motion.div
            initial={{ scale: 0.9, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            exit={{ scale: 0.9, opacity: 0 }}
            className="relative bg-white w-full max-w-lg rounded-3xl shadow-2xl overflow-hidden"
          >
            <div className="bg-green-800 p-6 text-white text-center">
              <h2 className="text-xl font-bold">Form Data Siswa</h2>
            </div>
            <form onSubmit={handleSaveSiswa} className="p-8 space-y-4">
              <div className="grid grid-cols-2 gap-4">
                <div className="col-span-2">
                  <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">
                    NISN
                  </label>
                  <input
                    type="text"
                    required
                    disabled={
                      !!(
                        editingSiswa &&
                        students.some(
                          (s) => s.nisn === editingSiswa.nisn && s.nisn !== "",
                        )
                      )
                    }
                    value={editingSiswa?.nisn || ""}
                    onChange={(e) =>
                      setEditingSiswa({
                        ...(editingSiswa as Student),
                        nisn: e.target.value,
                      })
                    }
                    className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold disabled:opacity-50 disabled:cursor-not-allowed"
                  />
                </div>
                <div>
                  <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">
                    Nama Lengkap
                  </label>
                  <input
                    type="text"
                    required
                    value={editingSiswa?.nama || ""}
                    onChange={(e) =>
                      setEditingSiswa({
                        ...(editingSiswa as Student),
                        nama: e.target.value,
                      })
                    }
                    className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold"
                  />
                </div>
                <div>
                  <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">
                    Jenis Kelamin
                  </label>
                  <select
                    required
                    value={editingSiswa?.jenisKelamin || ""}
                    onChange={(e) =>
                      setEditingSiswa({
                        ...(editingSiswa as Student),
                        jenisKelamin: e.target.value as "L" | "P",
                      })
                    }
                    className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold"
                  >
                    <option value="">Pilih Gender</option>
                    <option value="L">Laki-Laki (L)</option>
                    <option value="P">Perempuan (P)</option>
                  </select>
                </div>
                <div>
                  <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">
                    Kelas
                  </label>
                  <select
                    required
                    value={editingSiswa?.kelas || ""}
                    onChange={(e) =>
                      setEditingSiswa({
                        ...(editingSiswa as Student),
                        kelas: e.target.value,
                      })
                    }
                    className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold"
                  >
                    <option value="">Pilih</option>
                    {classrooms.map((k) => (
                      <option key={k.nama} value={k.nama}>
                        {k.nama}
                      </option>
                    ))}
                  </select>
                </div>
                <div className="md:grid md:grid-cols-2 gap-4">
                  <div>
                    <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">
                      Ayah
                    </label>
                    <input
                      type="text"
                      value={editingSiswa?.ayah || ""}
                      onChange={(e) =>
                        setEditingSiswa({
                          ...(editingSiswa as Student),
                          ayah: e.target.value,
                        })
                      }
                      className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold"
                      placeholder="Nama Ayah"
                    />
                  </div>
                  <div>
                    <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">
                      Ibu
                    </label>
                    <input
                      type="text"
                      value={editingSiswa?.ibu || ""}
                      onChange={(e) =>
                        setEditingSiswa({
                          ...(editingSiswa as Student),
                          ibu: e.target.value,
                        })
                      }
                      className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold"
                      placeholder="Nama Ibu"
                    />
                  </div>
                </div>
                <div>
                  <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">
                    HP Orang Tua
                  </label>
                  <input
                    type="text"
                    value={editingSiswa?.hp || ""}
                    onChange={(e) =>
                      setEditingSiswa({
                        ...(editingSiswa as Student),
                        hp: e.target.value,
                      })
                    }
                    className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold"
                  />
                </div>
                <div className="col-span-2">
                  <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">
                    Foto Siswa (Data URL)
                  </label>
                  <div className="flex items-center gap-4">
                    {editingSiswa?.foto && (
                      <img
                        src={editingSiswa.foto}
                        className="w-12 h-12 rounded-lg object-cover border"
                        alt="Profile"
                      />
                    )}
                    <input
                      type="file"
                      accept="image/*"
                      onChange={(e) => {
                        const file = e.target.files?.[0];
                        if (file) {
                          const reader = new FileReader();
                          reader.onloadend = () => {
                            setEditingSiswa({
                              ...(editingSiswa as Student),
                              foto: reader.result as string,
                            });
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
                <button
                  type="button"
                  onClick={() => setShowSiswaModal(false)}
                  className="flex-1 py-4 font-bold text-zinc-400"
                >
                  Batal
                </button>
                <button
                  type="submit"
                  className="flex-1 bg-green-800 text-white rounded-2xl py-4 font-bold shadow-xl"
                >
                  Simpan Data
                </button>
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
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="absolute inset-0 bg-black/60 backdrop-blur-sm"
            onClick={() => setShowGuruModal(false)}
          />
          <motion.div
            initial={{ scale: 0.9, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            exit={{ scale: 0.9, opacity: 0 }}
            className="relative bg-white w-full max-w-lg rounded-3xl shadow-2xl overflow-hidden"
          >
            <div className="bg-zinc-900 p-6 text-white text-center">
              <h2 className="text-xl font-bold">Form Data Guru</h2>
            </div>
            <form onSubmit={handleSaveGuru} className="p-8 space-y-4">
              <div className="grid grid-cols-2 gap-4">
                <div className="col-span-2">
                  <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">
                    NIP
                  </label>
                  <input
                    type="text"
                    required
                    disabled={
                      !!(
                        editingGuru &&
                        teachers.some(
                          (t) => t.nip === editingGuru.nip && t.nip !== "",
                        )
                      )
                    }
                    value={editingGuru?.nip || ""}
                    onChange={(e) =>
                      setEditingGuru({ ...editingGuru, nip: e.target.value })
                    }
                    className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold disabled:opacity-50 disabled:cursor-not-allowed"
                  />
                </div>
                <div className="col-span-2">
                  <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">
                    Nama Guru
                  </label>
                  <input
                    type="text"
                    required
                    value={editingGuru?.nama || ""}
                    onChange={(e) =>
                      setEditingGuru({ ...editingGuru, nama: e.target.value })
                    }
                    className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold"
                  />
                </div>
                <div>
                  <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">
                    Jabatan
                  </label>
                  <select
                    required
                    value={editingGuru?.jabatan || ""}
                    onChange={(e) => {
                      const jabatan = e.target.value as any;
                      setEditingGuru({
                        ...editingGuru,
                        jabatan,
                        kelas:
                          jabatan === "Guru Wali Kelas"
                            ? editingGuru.kelas
                            : "",
                      });
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
                {editingGuru?.jabatan === "Guru Wali Kelas" && (
                  <div>
                    <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">
                      Kelas Wali
                    </label>
                    <select
                      required
                      value={editingGuru?.kelas || ""}
                      onChange={(e) =>
                        setEditingGuru({
                          ...editingGuru,
                          kelas: e.target.value,
                        })
                      }
                      className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold"
                    >
                      <option value="">-- Pilih Kelas --</option>
                      {classrooms.map((c) => (
                        <option key={c.nama} value={c.nama}>
                          {c.nama}
                        </option>
                      ))}
                    </select>
                  </div>
                )}
                <div className="col-span-2">
                  <div className="grid grid-cols-2 gap-4">
                    <div>
                      <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">
                        Username Login
                      </label>
                      <input
                        type="text"
                        required
                        value={editingGuru?.user || ""}
                        onChange={(e) =>
                          setEditingGuru({
                            ...editingGuru,
                            user: e.target.value,
                          })
                        }
                        className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold"
                      />
                    </div>
                    <div>
                      <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">
                        Password
                      </label>
                      <input
                        type="text"
                        required
                        value={editingGuru?.pass || ""}
                        onChange={(e) =>
                          setEditingGuru({
                            ...editingGuru,
                            pass: e.target.value,
                          })
                        }
                        className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold"
                      />
                    </div>
                  </div>
                </div>
              </div>
              <div className="pt-4 flex gap-3">
                <button
                  type="button"
                  onClick={() => setShowGuruModal(false)}
                  className="flex-1 py-4 font-bold text-zinc-400"
                >
                  Batal
                </button>
                <button
                  type="submit"
                  className="flex-1 bg-zinc-900 text-white rounded-2xl py-4 font-bold shadow-xl"
                >
                  Simpan Akun
                </button>
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
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="absolute inset-0 bg-black/60 backdrop-blur-sm"
            onClick={() => setShowKelasModal(false)}
          />
          <motion.div
            initial={{ scale: 0.9, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            exit={{ scale: 0.9, opacity: 0 }}
            className="relative bg-white w-full max-w-md rounded-3xl shadow-2xl overflow-hidden"
          >
            <div className="bg-blue-900 p-6 text-white text-center">
              <h2 className="text-xl font-bold">Konfigurasi Ruang Kelas</h2>
            </div>
            <form onSubmit={handleSaveKelas} className="p-8 space-y-4">
              <div>
                <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">
                  Nama Kelas
                </label>
                <input
                  type="text"
                  required
                  disabled={
                    !!(
                      editingKelas &&
                      classrooms.some(
                        (c) =>
                          c.nama ===
                            (editingKelas.oldNama || editingKelas.nama) &&
                          c.nama !== "",
                      )
                    )
                  }
                  value={editingKelas?.nama || ""}
                  onChange={(e) =>
                    setEditingKelas({ ...editingKelas, nama: e.target.value })
                  }
                  className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold disabled:opacity-50 disabled:cursor-not-allowed"
                  placeholder="Contoh: VII-A"
                />
              </div>
              <div>
                <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">
                  Wali Kelas
                </label>
                <input
                  type="text"
                  value={editingKelas?.wali || ""}
                  onChange={(e) =>
                    setEditingKelas({ ...editingKelas, wali: e.target.value })
                  }
                  className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold"
                />
              </div>
              <div>
                <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">
                  Jumlah Siswa
                </label>
                <input
                  type="number"
                  value={editingKelas?.jumlah || ""}
                  onChange={(e) =>
                    setEditingKelas({
                      ...editingKelas,
                      jumlah: parseInt(e.target.value),
                    })
                  }
                  className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold"
                />
              </div>
              <div className="pt-4 flex gap-3">
                <button
                  type="button"
                  onClick={() => setShowKelasModal(false)}
                  className="flex-1 py-4 font-bold text-zinc-400"
                >
                  Tutup
                </button>
                <button
                  type="submit"
                  className="flex-1 bg-blue-900 text-white rounded-2xl py-4 font-bold shadow-xl"
                >
                  Terapkan
                </button>
              </div>
            </form>
          </motion.div>
        </div>
      )}
    </AnimatePresence>
  );

  const [teachingSessions, setTeachingSessions] = useState<
    { kelas: string; hari: string; target: number; jps?: number[] }[]
  >([]);
  const [openJpDropdown, setOpenJpDropdown] = useState<number | null>(null);

  const handleSaveJadwal = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingJadwal?.nip) return;
    toggleLoader(true);
    try {
      // Cleanup existing schedules for this specific teacher, mapel, and class
      // to avoid duplicates if days are changed.
      const teacherSchedules = teachingSchedules.filter(
        (ts) =>
          ts.nip === editingJadwal.nip &&
          ts.mapel === editingJadwal.mapel &&
          ts.kelas === editingJadwal.kelas,
      );
      for (const old of teacherSchedules) {
        await firestoreService.hapusJadwalMengajar(old.id);
      }

      const g = teachers.find((t) => t.nip === editingJadwal.nip);
      // Ensure all sessions have the correct class from editingJadwal
      const finalSessions = teachingSessions.map((s) => ({
        ...s,
        kelas: editingJadwal.kelas || "",
      }));

      await firestoreService.saveTeachingBatch(
        editingJadwal.nip,
        g?.nama || "",
        finalSessions,
        editingJadwal.mapel || "",
      );
      refreshMasterData();
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
    const link = document.createElement("a");
    link.href = dataUrl;
    link.download = fileName;
    link.click();
  };

  const renderProfilSiswa = () => {
    const teacher = teachers.find((t) => t.nip === session?.uid);
    const jabatan =
      teacher?.jabatan || (session?.role === "Admin" ? "Wakamad" : "Guru");
    const isKamadWakamad =
      ["Kamad", "Wakamad"].includes(jabatan) || session?.role === "Admin";

    const waliKelas = session?.kelas;

    // Determine which students to show
    let filteredStudents = [];
    let title = "";
    let description = "";

    if (isKamadWakamad) {
      filteredStudents = students.filter((s) =>
        studentProfileClassFilter
          ? s.kelas === studentProfileClassFilter
          : true,
      );
      title = studentProfileClassFilter
        ? `Profil Siswa Kelas ${studentProfileClassFilter}`
        : "Seluruh Profil Siswa";
      description = "Melihat profil siswa untuk seluruh tingkatan/kelas.";
    } else if (waliKelas) {
      filteredStudents = students.filter((s) => s.kelas === waliKelas);
      title = `Profil Siswa Kelas ${waliKelas}`;
      description = "Daftar siswa dalam perwakilan kelas Anda.";
    } else {
      return (
        <div className="p-8 text-center text-gray-400">
          Akses Dibatasi. Anda bukan Wali Kelas atau Wakamad.
        </div>
      );
    }

    return (
      <motion.div
        initial={{ opacity: 0 }}
        animate={{ opacity: 1 }}
        className="space-y-6 pb-20"
      >
        <div className="flex flex-col md:flex-row justify-between items-start md:items-center bg-white p-6 md:p-8 rounded-[2.5rem] shadow-sm border border-gray-100 gap-4">
          <div>
            <h2 className="text-xl font-black flex items-center gap-2 uppercase tracking-tight">
              <GraduationCap className="text-green-700" size={24} /> {title}
            </h2>
            <p className="text-xs text-gray-400 font-bold mt-1 uppercase tracking-widest">
              {description}
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-4 w-full md:w-auto">
            {isKamadWakamad && (
              <div className="flex items-center gap-2 bg-zinc-50 px-4 py-2 rounded-2xl border border-zinc-100">
                <TableIcon size={14} className="text-zinc-400" />
                <select
                  value={studentProfileClassFilter}
                  onChange={(e) => setStudentProfileClassFilter(e.target.value)}
                  className="bg-transparent border-0 text-xs font-black uppercase tracking-widest focus:ring-0 text-zinc-600"
                >
                  <option value="">Semua Kelas</option>
                  {classrooms.map((c) => (
                    <option key={c.nama} value={c.nama}>
                      {c.nama}
                    </option>
                  ))}
                </select>
              </div>
            )}
            <div className="text-right bg-green-50 px-6 py-2 rounded-2xl border border-green-100">
              <p className="text-[9px] font-black text-green-600 uppercase tracking-widest mb-1">
                Total Tampil
              </p>
              <p className="text-xl font-black text-green-800">
                {filteredStudents.length}
              </p>
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
                  <tr
                    key={idx}
                    className="hover:bg-zinc-50 transition-colors group"
                  >
                    <td className="px-8 py-5">
                      <div className="flex items-center gap-4">
                        <div className="w-10 h-10 rounded-2xl bg-zinc-100 flex items-center justify-center overflow-hidden border border-zinc-200 group-hover:border-green-300 transition-colors">
                          {s.foto ? (
                            <img
                              src={s.foto}
                              className="w-full h-full object-cover"
                            />
                          ) : (
                            <User size={18} className="text-zinc-300" />
                          )}
                        </div>
                        <div>
                          <span className="font-black text-zinc-900 block leading-tight">
                            {s.nama}
                          </span>
                          <span className="text-[10px] font-bold text-zinc-400 uppercase tracking-tighter">
                            Status Aktif
                          </span>
                        </div>
                      </div>
                    </td>
                    <td className="px-8 py-5">
                      <span className="font-bold text-zinc-500 block">
                        {s.nisn}
                      </span>
                      <span className="text-[10px] font-black text-green-700 bg-green-50 px-2 py-0.5 rounded-lg uppercase tracking-widest">
                        {s.kelas}
                      </span>
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
                        <GraduationCap
                          size={48}
                          className="mx-auto mb-4 opacity-20"
                        />
                        <p className="font-bold uppercase tracking-widest text-[10px]">
                          Data tidak ditemukan
                        </p>
                        <p className="text-xs font-medium mt-1">
                          Belum ada siswa yang terdaftar di pilihan ini.
                        </p>
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
      {selectedStudentCard && activePanel === "profil-siswa" && (
        <div className="fixed inset-0 z-[200] flex items-center justify-center p-4">
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="absolute inset-0 bg-black/60 backdrop-blur-sm"
            onClick={() => setSelectedStudentCard(null)}
          />
          <motion.div
            initial={{ scale: 0.9, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            exit={{ scale: 0.9, opacity: 0 }}
            className="relative bg-white w-full max-w-2xl rounded-[2.5rem] shadow-2xl overflow-hidden max-h-[90vh] flex flex-col"
          >
            <div className="p-8 border-b flex justify-between items-center bg-gray-50">
              <h2 className="text-xl font-black text-green-950 uppercase tracking-widest">
                Detail Profil Siswa
              </h2>
              <button
                onClick={() => setSelectedStudentCard(null)}
                className="text-gray-400 hover:text-red-500 transition-colors"
              >
                <XCircle size={28} />
              </button>
            </div>

            <div className="p-8 overflow-y-auto custom-scrollbar flex-1">
              <div className="flex flex-col md:flex-row gap-8 mb-10 items-center md:items-start">
                <div className="relative group">
                  <div className="w-40 h-40 rounded-[2.5rem] bg-gray-100 overflow-hidden border-4 border-white shadow-xl relative">
                    {selectedStudentCard.foto ? (
                      <img
                        src={selectedStudentCard.foto}
                        className="w-full h-full object-cover"
                        crossOrigin="anonymous"
                      />
                    ) : (
                      <div className="w-full h-full flex items-center justify-center text-gray-300">
                        <User size={64} />
                      </div>
                    )}
                  </div>
                  {selectedStudentCard.foto && (
                    <button
                      onClick={() =>
                        downloadPhoto(
                          selectedStudentCard.foto!,
                          `Foto-${selectedStudentCard.nama}.png`,
                        )
                      }
                      className="absolute -bottom-2 -right-2 bg-green-600 text-white p-3 rounded-2xl shadow-lg border-4 border-white hover:bg-green-700 transition-all scale-90 group-hover:scale-100"
                      title="Download Foto"
                    >
                      <Download size={18} />
                    </button>
                  )}
                </div>

                <div className="flex-1 text-center md:text-left">
                  <h3 className="text-2xl font-black text-green-900 mb-2">
                    {selectedStudentCard.nama}
                  </h3>
                  <p className="text-xs font-bold text-green-600 uppercase tracking-widest bg-green-50 px-4 py-1.5 rounded-full inline-block">
                    Kelas {selectedStudentCard.kelas}
                  </p>
                  <div className="mt-6 grid grid-cols-2 gap-4">
                    <div className="bg-gray-50 p-4 rounded-2xl border border-gray-100">
                      <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest mb-1">
                        Jenis Kelamin
                      </p>
                      <p className="font-bold text-sm">
                        {selectedStudentCard.jenisKelamin === "L"
                          ? "Laki-Laki"
                          : "Perempuan"}
                      </p>
                    </div>
                    <div className="bg-gray-50 p-4 rounded-2xl border border-gray-100">
                      <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest mb-1">
                        Tempat Lahir
                      </p>
                      <p className="font-bold text-sm">
                        {selectedStudentCard.tempat || "-"}
                      </p>
                    </div>
                    <div className="bg-gray-50 p-4 rounded-2xl border border-gray-100">
                      <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest mb-1">
                        Tanggal Lahir
                      </p>
                      <p className="font-bold text-sm">
                        {selectedStudentCard.tgl
                          ? formatIndoDate(selectedStudentCard.tgl)
                          : "-"}
                      </p>
                    </div>
                  </div>
                </div>
              </div>

              <div className="space-y-6">
                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  <div className="space-y-2">
                    <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest ml-1">
                      Nama Ayah
                    </label>
                    <input
                      type="text"
                      readOnly
                      value={selectedStudentCard.ayah || "-"}
                      className="w-full bg-gray-50 border-0 rounded-2xl px-6 py-4 text-sm font-bold text-gray-500 cursor-not-allowed"
                    />
                  </div>
                  <div className="space-y-2">
                    <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest ml-1">
                      Nama Ibu
                    </label>
                    <input
                      type="text"
                      readOnly
                      value={selectedStudentCard.ibu || "-"}
                      className="w-full bg-gray-50 border-0 rounded-2xl px-6 py-4 text-sm font-bold text-gray-500 cursor-not-allowed"
                    />
                  </div>
                </div>
                <div className="space-y-2">
                  <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest ml-1">
                    WhatsApp Orang Tua
                  </label>
                  <input
                    type="text"
                    readOnly
                    value={selectedStudentCard.hp || "-"}
                    className="w-full bg-gray-50 border-0 rounded-2xl px-6 py-4 text-sm font-bold text-gray-500 cursor-not-allowed"
                  />
                </div>
              </div>
            </div>

            <div className="p-8 border-t bg-gray-50 flex justify-center">
              <button
                onClick={() => setSelectedStudentCard(null)}
                className="bg-zinc-900 text-white px-12 py-4 rounded-2xl font-black text-xs uppercase tracking-widest hover:bg-zinc-800 transition-all shadow-xl"
              >
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
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="absolute inset-0 bg-black/60 backdrop-blur-sm"
            onClick={() => setShowJadwalModal(false)}
          />
          <motion.div
            initial={{ scale: 0.9, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            exit={{ scale: 0.9, opacity: 0 }}
            className="relative bg-white w-full max-w-lg rounded-3xl shadow-2xl overflow-hidden"
          >
            <div className="bg-green-950 p-6 text-white text-center">
              <h2 className="text-xl font-bold">Konfigurasi Jam Mengajar</h2>
            </div>
            <div className="p-8 space-y-6 max-h-[75vh] overflow-y-auto custom-scrollbar">
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div>
                  <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">
                    Nama Guru
                  </label>
                  <select
                    required
                    value={editingJadwal?.nip || ""}
                    onChange={(e) => {
                      const t = teachers.find(
                        (tg) => tg.nip === e.target.value,
                      );
                      setEditingJadwal({
                        ...(editingJadwal as any),
                        nip: e.target.value,
                        namaGuru: t?.nama || "",
                      });
                    }}
                    className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold mt-1 text-sm focus:ring-2 focus:ring-green-500"
                  >
                    <option value="">-- Pilih Guru --</option>
                    {teachers
                      .filter((t) => t.nip !== "ADMIN001" && t.role !== "Admin")
                      .map((t) => (
                        <option key={t.nip} value={t.nip}>
                          {t.nama}
                        </option>
                      ))}
                  </select>
                </div>
                <div>
                  <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">
                    Mata Pelajaran
                  </label>
                  <select
                    required
                    value={editingJadwal?.mapel || ""}
                    onChange={(e) =>
                      setEditingJadwal({
                        ...(editingJadwal as any),
                        mapel: e.target.value,
                      })
                    }
                    className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold mt-1 text-sm focus:ring-2 focus:ring-green-500"
                  >
                    <option value="">-- Pilih Mapel --</option>
                    {activeSubjects.map((s) => (
                      <option key={s} value={s}>
                        {s}
                      </option>
                    ))}
                  </select>
                </div>
                <div>
                  <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">
                    Kelas
                  </label>
                  <select
                    required
                    value={editingJadwal?.kelas || ""}
                    onChange={(e) =>
                      setEditingJadwal({
                        ...(editingJadwal as any),
                        kelas: e.target.value,
                      })
                    }
                    className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold mt-1 text-sm focus:ring-2 focus:ring-green-500"
                  >
                    <option value="">-- Pilih Kelas --</option>
                    {classrooms.map((c) => (
                      <option key={c.nama} value={c.nama}>
                        {c.nama}
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              <div>
                <div className="flex justify-between items-center mb-3">
                  <label className="text-[10px] font-black text-zinc-400 uppercase tracking-widest ml-1">
                    Sesi Mengajar (Hari & Target)
                  </label>
                  <button
                    onClick={() =>
                      setTeachingSessions([
                        ...teachingSessions,
                        {
                          kelas: editingJadwal?.kelas || "",
                          hari: "Senin",
                          target: 2,
                          jps: [],
                        },
                      ])
                    }
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
                    <div
                      key={idx}
                      className="bg-gray-50 p-4 rounded-2xl border border-gray-100 flex items-center gap-4 relative group"
                    >
                      <div className="flex-1">
                        <label className="text-[9px] font-black text-gray-400 uppercase mb-1 ml-1 block">
                          Hari
                        </label>
                        <select
                          value={session.hari}
                          onChange={(e) => {
                            const newSess = [...teachingSessions];
                            newSess[idx].hari = e.target.value;
                            newSess[idx].jps = []; // Reset selected JPs since active JPs might differ for this day
                            setTeachingSessions(newSess);
                          }}
                          className="w-full bg-white border border-gray-200 rounded-xl px-3 py-2 text-xs font-bold focus:ring-1 focus:ring-green-500"
                        >
                          {[
                            "Senin",
                            "Selasa",
                            "Rabu",
                            "Kamis",
                            "Jumat",
                            "Sabtu",
                          ].map((h) => (
                            <option key={h} value={h}>
                              {h}
                            </option>
                          ))}
                        </select>
                      </div>

                      {/* Multiselect Dropdown for JPs */}
                      <div className="w-44 relative">
                        <label className="text-[9px] font-black text-gray-400 uppercase mb-1 ml-1 block">
                          Jam Pelajaran (JP)
                        </label>
                        <div className="relative">
                          <button
                            type="button"
                            onClick={() =>
                              setOpenJpDropdown(
                                openJpDropdown === idx ? null : idx,
                              )
                            }
                            className="w-full bg-white border border-gray-200 rounded-xl px-3 py-2 text-xs font-bold text-left flex justify-between items-center gap-1 focus:ring-1 focus:ring-green-500 cursor-pointer min-h-[38px] truncate"
                          >
                            <span className="truncate">
                              {session.jps && session.jps.length > 0
                                ? session.jps.map((j) => `JP ${j}`).join(", ")
                                : "-- Pilih JP --"}
                            </span>
                            <span className="text-gray-400 text-[9px]">▼</span>
                          </button>

                          {openJpDropdown === idx &&
                            (() => {
                              const daySett = settings.find(
                                (s) => s.hari === session.hari,
                              );
                              const maxJp = session.hari === "Jumat" ? 6 : 8;
                              const activeList =
                                daySett && Array.isArray(daySett.activeJps)
                                  ? daySett.activeJps
                                  : Array.from(
                                      { length: maxJp },
                                      (_, i) => i + 1,
                                    );

                              return (
                                <div className="absolute left-0 right-0 mt-1 bg-white border border-gray-200 rounded-xl shadow-lg z-50 p-2 max-h-48 overflow-y-auto">
                                  {activeList.length === 0 ? (
                                    <p className="text-[10px] text-gray-400 p-2 italic text-center">
                                      Tidak ada JP aktif
                                    </p>
                                  ) : (
                                    <div className="space-y-1">
                                      {activeList.map((j) => {
                                        const isSelected = (
                                          session.jps || []
                                        ).includes(j);
                                        return (
                                          <label
                                            key={j}
                                            className="flex items-center gap-2 p-1.5 hover:bg-gray-50 rounded-lg text-xs font-bold cursor-pointer text-zinc-700"
                                          >
                                            <input
                                              type="checkbox"
                                              checked={isSelected}
                                              onChange={(e) => {
                                                const checked =
                                                  e.target.checked;
                                                const currentJps =
                                                  session.jps || [];
                                                let nextJps: number[];
                                                if (checked) {
                                                  nextJps = [
                                                    ...currentJps,
                                                    j,
                                                  ].sort((a, b) => a - b);
                                                } else {
                                                  nextJps = currentJps.filter(
                                                    (n) => n !== j,
                                                  );
                                                }
                                                const updated = [
                                                  ...teachingSessions,
                                                ];
                                                updated[idx] = {
                                                  ...session,
                                                  jps: nextJps,
                                                };
                                                setTeachingSessions(updated);
                                              }}
                                              className="w-3.5 h-3.5 text-green-600 border-gray-300 rounded focus:ring-green-500"
                                            />
                                            <span>JP {j}</span>
                                          </label>
                                        );
                                      })}
                                    </div>
                                  )}
                                </div>
                              );
                            })()}
                        </div>
                      </div>

                      <div className="w-24">
                        <label className="text-[9px] font-black text-gray-400 uppercase mb-1 ml-1 block">
                          Sesi/Jam
                        </label>
                        <input
                          type="number"
                          min={1}
                          value={session.target}
                          onChange={(e) => {
                            const newSess = [...teachingSessions];
                            newSess[idx].target = parseInt(e.target.value);
                            setTeachingSessions(newSess);
                          }}
                          className="w-full bg-white border border-gray-200 rounded-xl px-3 py-2 text-xs font-bold focus:ring-1 focus:ring-green-500"
                        />
                      </div>
                      <button
                        onClick={() =>
                          setTeachingSessions(
                            teachingSessions.filter((_, i) => i !== idx),
                          )
                        }
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
              <button
                type="button"
                onClick={() => setShowJadwalModal(false)}
                className="flex-1 py-4 font-bold text-zinc-400"
              >
                Batal
              </button>
              <button
                onClick={handleSaveJadwal}
                disabled={
                  !editingJadwal?.nip ||
                  !editingJadwal?.mapel ||
                  !editingJadwal?.kelas ||
                  teachingSessions.length === 0
                }
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
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="absolute inset-0 bg-black/60 backdrop-blur-sm"
            onClick={() => setShowSessionDetail(null)}
          />
          <motion.div
            initial={{ scale: 0.9, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            exit={{ scale: 0.9, opacity: 0 }}
            className="relative bg-white w-full max-w-sm rounded-[2.5rem] shadow-2xl overflow-hidden border border-gray-100"
          >
            <div className="bg-green-100 p-8 text-center">
              <div className="bg-white w-16 h-16 rounded-3xl flex items-center justify-center mx-auto mb-4 shadow-sm">
                <School className="text-green-800" size={32} />
              </div>
              <h2 className="text-lg font-black text-green-950 uppercase tracking-tight">
                {showSessionDetail.name}
              </h2>
              <p className="text-xs font-bold text-green-700/60 uppercase tracking-widest mt-1">
                Detail Jadwal Mengajar
              </p>
            </div>
            <div className="p-8">
              <div className="space-y-4">
                {showSessionDetail.mapping.map((s, idx) => (
                  <div
                    key={idx}
                    className="bg-gray-50 p-4 rounded-2xl border border-gray-100 flex justify-between items-center"
                  >
                    <div>
                      <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest mb-1">
                        {s.kelas}
                      </p>
                      <p className="font-bold text-sm text-gray-900">
                        {s.hari}
                      </p>
                    </div>
                    <div className="text-right">
                      <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest mb-1">
                        Target
                      </p>
                      <p className="font-black text-green-800 text-sm">
                        {s.targetPertemuan} Sesi
                      </p>
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

  const currentSiswaData = students.find((s) => s.nisn === session?.uid);

  const [biodataForm, setBiodataForm] = useState({
    nama: "",
    jenisKelamin: "",
    tempat: "",
    tgl: "",
    ayah: "",
    ibu: "",
    hp: "",
  });

  useEffect(() => {
    if (currentSiswaData) {
      setBiodataForm({
        nama: currentSiswaData.nama || "",
        jenisKelamin: currentSiswaData.jenisKelamin || "L",
        tempat: currentSiswaData.tempat || "",
        tgl: currentSiswaData.tgl || "",
        ayah: currentSiswaData.ayah || "",
        ibu: currentSiswaData.ibu || "",
        hp: currentSiswaData.hp || "",
      });
    }
  }, [currentSiswaData]);

  // Filtered lists for panels
  const filteredGuru = useMemo(() => {
    const list = teachers.filter(
      (t) => t.nip !== "ADMIN001" && t.role !== "Admin",
    );
    if (!searchTermGuru) return list;
    const s = searchTermGuru.toLowerCase();
    return list.filter(
      (t) =>
        t.nama.toLowerCase().includes(s) || t.nip.toLowerCase().includes(s),
    );
  }, [teachers, searchTermGuru]);

  const filteredSiswa = useMemo(() => {
    if (!searchTermSiswa) return students;
    const s = searchTermSiswa.toLowerCase();
    return students.filter(
      (st) =>
        st.nama.toLowerCase().includes(s) ||
        st.nisn.toLowerCase().includes(s) ||
        st.kelas.toLowerCase().includes(s),
    );
  }, [students, searchTermSiswa]);

  const filteredKelas = useMemo(() => {
    if (!searchTermKelas) return classrooms;
    const s = searchTermKelas.toLowerCase();
    return classrooms.filter(
      (c) =>
        c.nama.toLowerCase().includes(s) || c.wali.toLowerCase().includes(s),
    );
  }, [classrooms, searchTermKelas]);

  const filteredJadwal = useMemo(() => {
    let result = teachingSchedules;
    if (filterJadwalClass) {
      result = result.filter((ts) => ts.kelas === filterJadwalClass);
    }
    if (searchTermJadwal) {
      const s = searchTermJadwal.toLowerCase();
      result = result.filter(
        (ts) =>
          ts.namaGuru.toLowerCase().includes(s) ||
          ts.mapel.toLowerCase().includes(s) ||
          ts.kelas.toLowerCase().includes(s),
      );
    }
    return result;
  }, [teachingSchedules, searchTermJadwal, filterJadwalClass]);

  const filteredTeacherAttendance = useMemo(() => {
    return teacherAttendance.filter((ta) => {
      const matchesClass = filterGuruAbsensiClass
        ? ta.kelas === filterGuruAbsensiClass
        : true;
      const matchesDate = filterTanggalAbsensiGuru
        ? ta.tanggal === filterTanggalAbsensiGuru
        : true;
      return matchesClass && matchesDate;
    });
  }, [teacherAttendance, filterGuruAbsensiClass, filterTanggalAbsensiGuru]);

  const groupedJadwal = useMemo(() => {
    const groups: any = {};
    filteredJadwal.forEach((s) => {
      const key = `${s.nip}-${s.mapel || "N/A"}`;
      if (!groups[key]) {
        groups[key] = {
          nip: s.nip,
          namaGuru: s.namaGuru,
          mapel: s.mapel,
          sessions: [],
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
            <p className="text-[10px] font-black uppercase tracking-widest text-zinc-400 leading-none mb-1">
              {showSuccessToast.title}
            </p>
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

  const renderWarningToast = () => (
    <AnimatePresence>
      {showWarningToast && (
        <motion.div
          initial={{ opacity: 0, y: 100, scale: 0.8 }}
          animate={{ opacity: 1, y: 0, scale: 1 }}
          exit={{ opacity: 0, y: 20, scale: 0.8 }}
          className="fixed bottom-10 left-1/2 -translate-x-1/2 z-[310] bg-zinc-900 border border-zinc-800 text-white px-8 py-4 rounded-[2rem] shadow-2xl flex items-center gap-4 min-w-[320px] max-w-[90vw]"
        >
          <div className="bg-amber-500 p-2 rounded-xl text-white">
            <AlertTriangle size={24} />
          </div>
          <div className="min-w-0 flex-1">
            <p className="text-[10px] font-black uppercase tracking-widest text-zinc-400 leading-none mb-1 text-amber-500">
              {showWarningToast.title}
            </p>
            <p className="text-xs font-bold text-gray-100">{showWarningToast.message}</p>
          </div>
          <div className="absolute bottom-0 left-0 right-0 h-1 overflow-hidden rounded-b-[2rem]">
            <motion.div
              initial={{ width: "100%" }}
              animate={{ width: "0%" }}
              transition={{ duration: 4, ease: "linear" }}
              className="h-full bg-amber-500"
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
            <RefreshCcw
              className="animate-spin text-green-700 mb-4"
              size={48}
            />
            <p className="font-bold text-green-900">
              {loaderText || "Sinkronisasi Data..."}
            </p>
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
              initial={{ x: "-100%" }}
              animate={{ x: 0 }}
              exit={{ x: "-100%" }}
              transition={{ type: "spring", damping: 25, stiffness: 200 }}
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
                        target.style.opacity = "0";
                        const parent = target.parentElement;
                        if (parent) {
                          const icon = parent.querySelector(".fallback-icon");
                          if (icon) (icon as HTMLElement).style.opacity = "1";
                        }
                      } else {
                        target.src = defaultLogo;
                      }
                    }}
                  />
                  <School
                    size={36}
                    className="text-green-700 absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 opacity-0 fallback-icon transition-opacity"
                  />
                </div>
                <h2 className="font-black text-lg text-white">SIGAP MTsN 2</h2>
                <span className="bg-orange-500 text-black text-[10px] px-3 py-1 rounded-full font-black mt-2 uppercase tracking-widest">
                  {session?.jabatan || session?.role}
                </span>
              </div>

              <nav className="flex-grow py-6 overflow-y-auto">
                {menuItems.map((item) => (
                  <button
                    key={item.id}
                    onClick={() => {
                      setActivePanel(item.id);
                      setIsMobileMenuOpen(false);
                    }}
                    className={`w-full flex items-center px-8 py-4 transition-all text-sm font-bold border-l-4 ${activePanel === item.id ? "bg-green-900/20 text-green-400 border-green-500" : "text-zinc-500 border-transparent"}`}
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
        {menuItems.slice(0, 4).map((item) => (
          <button
            key={item.id}
            onClick={() => setActivePanel(item.id)}
            className={`flex flex-col items-center p-2 rounded-xl transition-all ${activePanel === item.id ? "text-green-500" : "text-zinc-500"}`}
          >
            <item.icon size={20} />
            <span className="text-[8px] font-bold mt-1 uppercase tracking-tighter truncate w-12 text-center">
              {item.id === "guru"
                ? "Guru"
                : item.id === "siswa"
                  ? "Siswa"
                  : item.label.split(" ")[0]}
            </span>
          </button>
        ))}
        {menuItems.length > 4 && (
          <div className="relative">
            <button
              onClick={() => setMobileExtraMenuOpen(!mobileExtraMenuOpen)}
              className={`flex flex-col items-center p-2 transition-all ${mobileExtraMenuOpen ? "text-green-500" : "text-zinc-500"}`}
            >
              <TableIcon size={20} />
              <span className="text-[8px] font-bold mt-1 uppercase tracking-tighter">
                Lainnya
              </span>
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
                    {menuItems.slice(4).map((item) => (
                      <button
                        key={item.id}
                        onClick={() => {
                          setActivePanel(item.id);
                          setMobileExtraMenuOpen(false);
                        }}
                        className={`w-full flex items-center px-4 py-4 text-xs font-bold rounded-xl transition-all ${activePanel === item.id ? "bg-green-900 text-white" : "text-zinc-400 hover:bg-zinc-800"}`}
                      >
                        <item.icon size={16} className="mr-3 text-green-500" />
                        {item.label}
                      </button>
                    ))}
                    <div className="border-t border-zinc-800 my-2 pt-2">
                      <button
                        onClick={() => {
                          handleLogout();
                          setMobileExtraMenuOpen(false);
                        }}
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

      {renderNaikKelasModal()}
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
              <h1 className="text-xl font-extrabold text-green-900">
                {menuItems.find((m) => m.id === activePanel)?.label ||
                  "Dashboard"}
              </h1>
              <button
                className={`p-2 hover:bg-gray-100 rounded-xl transition-all ${refreshing ? "text-green-600" : "text-gray-400"}`}
                onClick={() => triggerCurrentPanelRefresh()}
              >
                <RefreshCcw
                  size={16}
                  className={`${refreshing ? "animate-spin" : ""}`}
                />
              </button>
            </div>
          </div>
          <div className="flex items-center gap-4">
            <div className="text-right hidden sm:block">
              <p className="text-xs font-bold text-gray-400 uppercase tracking-widest">
                Selamat Datang
              </p>
              <p className="font-bold text-green-800">
                {session?.name || "User"}
              </p>
            </div>
            <div className="w-10 h-10 bg-green-100 rounded-full flex items-center justify-center text-green-700 font-bold border-2 border-green-200">
              {(session?.name || "U")[0]}
            </div>
          </div>
        </header>

        <AnimatePresence mode="wait">
          {activePanel === "home" && (
            <motion.div
              key="home"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              className="space-y-8"
            >
              {/* Today's Date Banner */}
              <div className="bg-white p-6 rounded-3xl border border-gray-100 shadow-sm flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 relative overflow-hidden group">
                <div className="absolute top-0 right-0 w-48 h-48 bg-green-50 rounded-full blur-3xl opacity-30 -mr-12 -mt-12 transition-transform group-hover:scale-110 duration-500"></div>
                <div className="flex items-center gap-4 relative z-10">
                  <div className="w-12 h-12 bg-green-50 text-green-700 rounded-2xl flex items-center justify-center border border-green-100 group-hover:rotate-6 transition-transform">
                    <Calendar size={24} />
                  </div>
                  <div>
                    <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest leading-none">
                      Hari Ini / Tanggal
                    </p>
                    <h2 className="text-xl sm:text-2xl font-black text-green-950 mt-1 whitespace-nowrap">
                      {new Date().toLocaleDateString("id-ID", {
                        weekday: "long",
                        day: "numeric",
                        month: "long",
                        year: "numeric",
                      })}
                    </h2>
                  </div>
                </div>
                <div className="px-4 py-2 bg-green-50 text-green-800 rounded-2xl text-[10px] font-black uppercase tracking-widest relative z-10 flex items-center gap-2 border border-green-100">
                  <span className="w-2.5 h-2.5 rounded-full bg-green-500 animate-pulse"></span>
                  Layanan Aktif / Real-time
                </div>
              </div>

              {/* Stats Grid */}
              <div className="grid grid-cols-2 lg:grid-cols-4 xl:grid-cols-7 gap-4">
                {[
                  {
                    label: "Total Siswa",
                    value: stats.siswaCount,
                    icon: GraduationCap,
                    color: "bg-blue-50 text-blue-600",
                    sub: `L: ${stats.siswaL} / P: ${stats.siswaP}`,
                  },
                  {
                    label: "Total Guru",
                    value: stats.guruCount,
                    icon: UserPlus,
                    color: "bg-zinc-100 text-zinc-600",
                  },
                  {
                    label: "Hadir",
                    value: stats.hadirHariIni,
                    icon: ClipboardList,
                    color: "bg-green-50 text-green-600",
                  },
                  {
                    label: "Terlambat",
                    value: stats.terlambatHariIni,
                    icon: Clock,
                    color: "bg-red-50 text-red-600",
                  },
                  {
                    label: "Sakit",
                    value: stats.sakitCount,
                    icon: Info,
                    color: "bg-amber-50 text-amber-600",
                  },
                  {
                    label: "Izin",
                    value: stats.izinCount,
                    icon: Info,
                    color: "bg-indigo-50 text-indigo-600",
                  },
                  {
                    label: "Alfa",
                    value: stats.alfaCount,
                    icon: XCircle,
                    color: "bg-rose-50 text-rose-600",
                  },
                ].map((s, i) => (
                  <div
                    key={i}
                    className="bg-white p-6 rounded-2xl shadow-sm border border-gray-100 relative overflow-hidden group"
                  >
                    <div
                      className={`${s.color} w-10 h-10 rounded-xl flex items-center justify-center mb-3 group-hover:scale-110 transition-transform`}
                    >
                      <s.icon size={20} />
                    </div>
                    <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest">
                      {s.label}
                    </p>
                    <div className="flex items-baseline gap-2">
                      <h3 className="text-2xl font-black mt-1">{s.value}</h3>
                      {s.sub && (
                        <span className="text-[9px] font-bold text-gray-400 uppercase tracking-tighter">
                          {s.sub}
                        </span>
                      )}
                    </div>
                  </div>
                ))}
              </div>

              {/* Quick Menu for Mobile */}
              <div className="md:hidden space-y-4">
                <div className="flex items-center justify-between">
                  <h3 className="text-xs font-black text-green-900 uppercase tracking-widest">
                    Menu Utama
                  </h3>
                </div>
                <div className="grid grid-cols-3 gap-3">
                  {menuItems.map((item) => (
                    <button
                      key={item.id}
                      onClick={() => setActivePanel(item.id)}
                      className={`bg-white p-4 rounded-2xl border border-gray-100 shadow-sm flex flex-col items-center justify-center transition-all ${activePanel === item.id ? "ring-2 ring-green-600" : ""}`}
                    >
                      <div className="bg-green-50 p-2 rounded-xl text-green-700 mb-2">
                        <item.icon size={20} />
                      </div>
                      <span className="text-[9px] font-bold text-gray-600 text-center leading-tight">
                        {item.label}
                      </span>
                    </button>
                  ))}
                </div>
              </div>

              {/* Today's Teacher Schedule OR Guru yang Mengajar Hari Ini for Kamad */}
              {session?.role === "Guru" && (
                <div className="bg-white p-8 rounded-[2rem] shadow-sm border border-gray-100 animate-fadeIn space-y-6">
                  {session?.jabatan === "Kamad" ? (
                    // Kamad Special Dashboard View: Guru yang Mengajar Hari Ini
                    <>
                      <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 border-b border-gray-50 pb-5">
                        <div>
                          <span className="text-[9px] font-black text-green-700 uppercase tracking-widest bg-green-50 px-2.5 py-1 rounded-full border border-green-100">
                            Live KBM Monitor
                          </span>
                          <h3 className="text-xl font-black text-zinc-900 mt-2">
                            Guru yang Mengajar Hari Ini
                          </h3>
                          <p className="text-xs text-gray-400 font-bold mt-1">
                            Gunakan panel monitoring ini untuk melihat guru yang
                            mengajar serta status absensinya.
                          </p>
                        </div>
                        <div className="flex items-center gap-3 bg-zinc-950 text-white px-5 py-3 rounded-2xl shadow-md shrink-0">
                          <Clock
                            size={16}
                            className="text-green-400 animate-pulse"
                          />
                          <div className="flex flex-col text-left">
                            <span className="text-[8px] font-black text-gray-400 uppercase tracking-widest">
                              Waktu Presensi
                            </span>
                            <span className="text-sm font-black font-mono tracking-wider">
                              {currentTime || "00:00"} WITA
                            </span>
                          </div>
                        </div>
                      </div>

                      {(() => {
                        const IndonesianDays = [
                          "Minggu",
                          "Senin",
                          "Selasa",
                          "Rabu",
                          "Kamis",
                          "Jumat",
                          "Sabtu",
                        ];
                        const todayDayName =
                          IndonesianDays[new Date().getDay()];

                        if (todayDayName === "Minggu") {
                          return (
                            <div className="py-8 text-center text-gray-400 italic font-medium">
                              Hari Minggu tidak ada jadwal KBM. Selamat
                              berlibur!
                            </div>
                          );
                        }

                        const todayDateStr = new Date().toLocaleDateString(
                          "en-CA",
                        );
                        const todaySetting = settings.find(
                          (st) => st.hari === todayDayName,
                        );
                        const currentHoliday = holidays.find(
                          (h) => h.tanggal === todayDateStr,
                        );

                        // Filter only real teachers (not students, not Kamad themselves, not Administrator)
                        const filteredTeachers = teachers.filter(
                          (t) =>
                            t.role !== "Siswa" &&
                            t.nip !== session?.uid &&
                            t.nip !== "ADMIN001" &&
                            t.role !== "Admin",
                        );

                        if (filteredTeachers.length === 0) {
                          return (
                            <div className="py-8 text-center text-gray-400 italic font-bold uppercase tracking-wider text-xs">
                              Belum ada Guru yang terdaftar di sistem.
                            </div>
                          );
                        }

                        const teachersTeachingToday = filteredTeachers.filter(
                          (t) =>
                            teachingSchedules.some(
                              (s) => s.nip === t.nip && s.hari === todayDayName,
                            ),
                        );
                        const teachersNotTeachingToday =
                          filteredTeachers.filter(
                            (t) =>
                              !teachingSchedules.some(
                                (s) =>
                                  s.nip === t.nip && s.hari === todayDayName,
                              ),
                          );

                        return (
                          <div className="space-y-6">
                            {/* Summary Cards */}
                            <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                              <div className="bg-emerald-50/50 border border-emerald-100 p-4 rounded-2xl flex items-center gap-4 shadow-xs">
                                <div className="w-12 h-12 rounded-xl bg-emerald-500 flex items-center justify-center text-white shrink-0">
                                  <BookOpen size={24} />
                                </div>
                                <div className="text-left">
                                  <p className="text-[10px] font-black text-emerald-800 uppercase tracking-widest leading-none">
                                    Total Guru Mengajar Hari Ini
                                  </p>
                                  <p className="text-2xl font-black text-emerald-950 mt-1">
                                    {teachersTeachingToday.length}{" "}
                                    <span className="text-xs font-bold text-emerald-600">
                                      Guru
                                    </span>
                                  </p>
                                </div>
                              </div>

                              <div className="bg-zinc-50 border border-zinc-200/80 p-4 rounded-2xl flex items-center gap-4 shadow-xs">
                                <div className="w-12 h-12 rounded-xl bg-zinc-400 flex items-center justify-center text-white shrink-0">
                                  <User size={24} />
                                </div>
                                <div className="text-left">
                                  <p className="text-[10px] font-black text-zinc-600 uppercase tracking-widest leading-none">
                                    Guru Tidak Mengajar Hari Ini
                                  </p>
                                  <p className="text-2xl font-black text-zinc-950 mt-1">
                                    {teachersNotTeachingToday.length}{" "}
                                    <span className="text-xs font-bold text-zinc-500">
                                      Guru
                                    </span>
                                  </p>
                                </div>
                              </div>
                            </div>

                            {/* Teachers Grid */}
                            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
                              {filteredTeachers.map((teacher) => {
                                // Check if teacher has teaching schedule for today
                                const teacherSchedulesToday =
                                  teachingSchedules.filter(
                                    (s) =>
                                      s.nip === teacher.nip &&
                                      s.hari === todayDayName,
                                  );
                                const isTeachingToday =
                                  teacherSchedulesToday.length > 0;

                                const limitJp =
                                  todayDayName === "Jumat" ? 6 : 8;

                                const getJpStartEnd = (
                                  jp: number,
                                  jpTimesFromDb: any,
                                  limit: number,
                                ) => {
                                  if (
                                    jpTimesFromDb?.[jp]?.start &&
                                    jpTimesFromDb?.[jp]?.end
                                  ) {
                                    return {
                                      start: jpTimesFromDb[jp].start,
                                      end: jpTimesFromDb[jp].end,
                                    };
                                  }
                                  // Fallback to standard calculated JP times:
                                  // Start KBM at 07:15. Each JP is 40 minutes.
                                  // After JP 4 (i.e., before JP 5 is started), there's a 20-min break.
                                  let startHour = 7;
                                  let startMin = 15;
                                  for (let i = 1; i < jp; i++) {
                                    if (i === 4) {
                                      startMin += 20; // 20 min break after JP 4
                                    }
                                    startMin += 40; // Add previous JP duration
                                  }
                                  // Normalize start
                                  startHour += Math.floor(startMin / 60);
                                  startMin = startMin % 60;

                                  let endHour = startHour;
                                  let endMin = startMin + 40;
                                  endHour += Math.floor(endMin / 60);
                                  endMin = endMin % 60;

                                  return {
                                    start: `${String(startHour).padStart(2, "0")}:${String(startMin).padStart(2, "0")}`,
                                    end: `${String(endHour).padStart(2, "0")}:${String(endMin).padStart(2, "0")}`,
                                  };
                                };

                                const parseTimeToMinutes = (tStr: string) => {
                                  if (!tStr) return 0;
                                  const parts = tStr.split(":");
                                  if (parts.length < 2) return 0;
                                  return (
                                    parseInt(parts[0]) * 60 + parseInt(parts[1])
                                  );
                                };

                                const currentMin = parseTimeToMinutes(
                                  currentTime || "00:00",
                                );

                                // Collect all classes & JP ranges where they teach today with exact active states
                                const teachDetailList =
                                  teacherSchedulesToday.map((sched) => {
                                    const schedJps = sched.jps || [];
                                    const minJp =
                                      schedJps.length > 0
                                        ? Math.min(...schedJps)
                                        : 1;
                                    const maxJp =
                                      schedJps.length > 0
                                        ? Math.max(...schedJps)
                                        : 1;

                                    const startJpTimes = getJpStartEnd(
                                      minJp,
                                      todaySetting?.jpTimes,
                                      limitJp,
                                    );
                                    const endJpTimes = getJpStartEnd(
                                      maxJp,
                                      todaySetting?.jpTimes,
                                      limitJp,
                                    );

                                    const startTimeStr = startJpTimes.start;
                                    const endTimeStr = endJpTimes.end;

                                    const startMin =
                                      parseTimeToMinutes(startTimeStr);
                                    const endMin =
                                      parseTimeToMinutes(endTimeStr);

                                    const attendanceRecord =
                                      teacherAttendance.find(
                                        (ta) =>
                                          ta.nip === teacher.nip &&
                                          ta.kelas === sched.kelas &&
                                          ta.tanggal === todayDateStr,
                                      );

                                    const scanned = !!attendanceRecord;
                                    const scanTime =
                                      attendanceRecord?.jam || "";

                                    // Check if this JP is inactive in today's settings
                                    const activeJpsList =
                                      todaySetting?.activeJps &&
                                      Array.isArray(todaySetting.activeJps)
                                        ? todaySetting.activeJps
                                        : Array.from(
                                            {
                                              length:
                                                todayDayName === "Jumat"
                                                  ? 6
                                                  : 8,
                                            },
                                            (_, i) => i + 1,
                                          );

                                    const isJpInactive = schedJps.some(
                                      (jp) => !activeJpsList.includes(jp),
                                    );

                                    let schedState:
                                      | "PASSED"
                                      | "ACTIVE"
                                      | "STANDBY"
                                      | "INACTIVE" = "STANDBY";
                                    if (isJpInactive) {
                                      schedState = "INACTIVE";
                                    } else if (currentMin < startMin) {
                                      schedState = "STANDBY";
                                    } else if (
                                      currentMin >= startMin &&
                                      currentMin <= endMin
                                    ) {
                                      schedState = "ACTIVE";
                                    } else {
                                      schedState = "PASSED";
                                    }

                                    return {
                                      kelas: sched.kelas,
                                      mapel: sched.mapel || "Mata Pelajaran",
                                      jps: schedJps,
                                      startTime: startTimeStr,
                                      endTime: endTimeStr,
                                      scanned,
                                      scanTime,
                                      state: schedState,
                                      reasonInactive:
                                        todaySetting?.reasonInactive || "",
                                    };
                                  });

                                // Identify combined teacher status
                                let teacherStatus:
                                  | "HIJAU"
                                  | "MERAH"
                                  | "STANDBY"
                                  | "ABU-ABU" = "ABU-ABU";
                                let statusLabel = "Tidak Mengajar";

                                const todayStatusRecord =
                                  teacherAttendance.find(
                                    (ta) =>
                                      ta.nip === teacher.nip &&
                                      ta.tanggal === todayDateStr &&
                                      ta.status,
                                  );

                                if (todayStatusRecord) {
                                  teacherStatus = "ABU-ABU";
                                  statusLabel = todayStatusRecord.status; // 'Sakit' or 'Izin' or 'Alfa'
                                } else if (currentHoliday) {
                                  if (isTeachingToday) {
                                    teacherStatus = "ABU-ABU";
                                    statusLabel = `Hari Libur ${currentHoliday.keterangan}`;
                                  } else {
                                    teacherStatus = "ABU-ABU";
                                    statusLabel = "Tidak Mengajar";
                                  }
                                } else if (isTeachingToday) {
                                  const nonInactiveDetails =
                                    teachDetailList.filter(
                                      (dt) => dt.state !== "INACTIVE",
                                    );

                                  if (
                                    teachDetailList.length > 0 &&
                                    nonInactiveDetails.length === 0
                                  ) {
                                    teacherStatus = "ABU-ABU";
                                    statusLabel = todaySetting?.reasonInactive
                                      ? `${todaySetting.reasonInactive}`
                                      : "Mengikuti Kegiatan";
                                  } else {
                                    const quietActive = nonInactiveDetails.some(
                                      (dt) => dt.state === "ACTIVE",
                                    );
                                    const quietStandby =
                                      nonInactiveDetails.some(
                                        (dt) => dt.state === "STANDBY",
                                      );
                                    const quietAllPassed =
                                      nonInactiveDetails.length > 0 &&
                                      nonInactiveDetails.every(
                                        (dt) => dt.state === "PASSED",
                                      );

                                    if (quietAllPassed) {
                                      teacherStatus = "ABU-ABU";
                                      statusLabel = "Selesai Mengajar";
                                    } else if (quietActive) {
                                      const activeClasses =
                                        nonInactiveDetails.filter(
                                          (dt) => dt.state === "ACTIVE",
                                        );
                                      const anyUnscannedActive =
                                        activeClasses.some((dt) => !dt.scanned);

                                      if (anyUnscannedActive) {
                                        teacherStatus = "MERAH";
                                        statusLabel = "Belum Scan";
                                      } else {
                                        teacherStatus = "HIJAU";
                                        statusLabel = "Aktif Mengajar";
                                      }
                                    } else if (quietStandby) {
                                      teacherStatus = "STANDBY";
                                      statusLabel = "Standby";
                                    } else {
                                      teacherStatus = "ABU-ABU";
                                      statusLabel = "Tidak Mengajar";
                                    }
                                  }
                                }

                                // Visual Styling Setup
                                let cardClass = "";
                                let circleClass = "";
                                let nameColorClass = "";
                                let roleColorClass = "";
                                let avatarClass = "";
                                let badgeClass = "";

                                switch (teacherStatus) {
                                  case "HIJAU":
                                    cardClass =
                                      "bg-emerald-50/15 border-emerald-200 hover:bg-emerald-50/35 shadow-xs";
                                    circleClass =
                                      "bg-green-500 border-green-600 shadow shadow-green-200";
                                    nameColorClass =
                                      "text-emerald-950 font-black";
                                    roleColorClass =
                                      "text-emerald-600/80 font-bold";
                                    avatarClass =
                                      "bg-green-100 text-green-700 border-green-200 animate-pulse";
                                    badgeClass =
                                      "bg-emerald-100 text-emerald-800 border-emerald-200 font-extrabold";
                                    break;
                                  case "MERAH":
                                    cardClass =
                                      "bg-rose-50/30 border-rose-300 hover:bg-rose-50/50 shadow-md ring-2 ring-rose-100 animate-pulse";
                                    circleClass =
                                      "bg-rose-500 border-rose-600 shadow shadow-rose-200";
                                    nameColorClass = "text-rose-950 font-black";
                                    roleColorClass =
                                      "text-rose-600/80 font-bold";
                                    avatarClass =
                                      "bg-rose-100 text-rose-700 border-rose-200";
                                    badgeClass =
                                      "bg-rose-100 text-rose-800 border-rose-200/90 font-extrabold";
                                    break;
                                  case "STANDBY":
                                    cardClass =
                                      "bg-amber-50/10 border-amber-200 hover:bg-amber-50/20";
                                    circleClass =
                                      "bg-amber-500 border-amber-600 shadow shadow-amber-200";
                                    nameColorClass =
                                      "text-amber-950 font-medium";
                                    roleColorClass =
                                      "text-amber-700 font-semibold";
                                    avatarClass =
                                      "bg-amber-100 text-amber-700 border-amber-150";
                                    badgeClass =
                                      "bg-amber-100 text-amber-800 border-amber-200 font-extrabold";
                                    break;
                                  case "ABU-ABU":
                                  default:
                                    cardClass =
                                      "bg-zinc-50/40 border-zinc-150 opacity-60 hover:opacity-100 transition-opacity";
                                    circleClass = "bg-zinc-300 border-zinc-400";
                                    nameColorClass =
                                      "text-zinc-650 font-semibold";
                                    roleColorClass =
                                      "text-zinc-400 font-medium";
                                    avatarClass =
                                      "bg-zinc-100 text-zinc-500 border-zinc-200";
                                    badgeClass =
                                      "bg-zinc-100 text-zinc-600 border-zinc-200 font-bold";
                                    break;
                                }

                                if (todayStatusRecord) {
                                  if (todayStatusRecord.status === "Sakit") {
                                    cardClass =
                                      "bg-orange-50/20 border-orange-200 hover:bg-orange-50/40 ring-1 ring-orange-100 shadow-sm opacity-100";
                                    circleClass =
                                      "bg-orange-500 border-orange-600 shadow shadow-orange-100";
                                    nameColorClass =
                                      "text-orange-950 font-black";
                                    roleColorClass =
                                      "text-orange-700/80 font-bold";
                                    avatarClass =
                                      "bg-orange-100 text-orange-700 border-orange-200";
                                    badgeClass =
                                      "bg-orange-100 text-orange-850 border-orange-200 font-extrabold";
                                  } else if (
                                    todayStatusRecord.status === "Izin"
                                  ) {
                                    cardClass =
                                      "bg-amber-50/20 border-amber-200 hover:bg-amber-50/40 ring-1 ring-amber-100 shadow-sm opacity-100";
                                    circleClass =
                                      "bg-amber-500 border-amber-600 shadow shadow-amber-150";
                                    nameColorClass =
                                      "text-amber-950 font-black";
                                    roleColorClass =
                                      "text-amber-700/80 font-bold";
                                    avatarClass =
                                      "bg-amber-100 text-amber-700 border-amber-200";
                                    badgeClass =
                                      "bg-amber-100 text-amber-850 border-amber-200 font-extrabold";
                                  } else if (
                                    todayStatusRecord.status === "Alfa" ||
                                    todayStatusRecord.status === "Alpa"
                                  ) {
                                    cardClass =
                                      "bg-rose-50/20 border-rose-200 hover:bg-rose-50/40 ring-1 ring-rose-100 shadow-sm opacity-100";
                                    circleClass =
                                      "bg-rose-500 border-rose-600 shadow shadow-rose-150";
                                    nameColorClass = "text-rose-950 font-black";
                                    roleColorClass =
                                      "text-rose-700/80 font-bold";
                                    avatarClass =
                                      "bg-rose-100 text-rose-700 border-rose-200";
                                    badgeClass =
                                      "bg-rose-100 text-rose-850 border-rose-200 font-extrabold";
                                  }
                                }

                                return (
                                  <div
                                    key={teacher.nip}
                                    onClick={() => {
                                      setExpandedKamadTeacherDetails(
                                        (prev) => ({
                                          ...prev,
                                          [teacher.nip]: !prev[teacher.nip],
                                        }),
                                      );
                                    }}
                                    className={`p-5 rounded-3xl border transition-all flex flex-col justify-between h-full relative overflow-hidden group cursor-pointer ${cardClass}`}
                                  >
                                    <div>
                                      <div className="flex justify-between items-start mb-4 gap-2">
                                        <div className="flex gap-3 items-center text-left">
                                          <div
                                            className={`w-10 h-10 rounded-full flex items-center justify-center font-black text-xs shrink-0 border ${avatarClass}`}
                                          >
                                            {teacher.nama
                                              ? teacher.nama
                                                  .split(" ")
                                                  .map((n: string) => n[0])
                                                  .slice(0, 2)
                                                  .join("")
                                                  .toUpperCase()
                                              : "??"}
                                          </div>
                                          <div className="flex flex-col">
                                            <span
                                              className={`text-xs font-black break-words line-clamp-1 group-hover:text-green-950 transition-colors uppercase tracking-tight ${nameColorClass}`}
                                            >
                                              {teacher.nama}
                                            </span>
                                            <span
                                              className={`text-[9px] uppercase tracking-widest ${roleColorClass}`}
                                            >
                                              {teacher.jabatan ||
                                                "Guru Pengajar"}
                                            </span>
                                          </div>
                                        </div>
                                        <div
                                          className={`w-3.5 h-3.5 rounded-full border flex items-center justify-center shrink-0 ${circleClass}`}
                                        >
                                          <span
                                            className="w-1.1 h-1.1 rounded-full bg-white animate-ping"
                                            style={{
                                              display:
                                                teacherStatus === "HIJAU" ||
                                                teacherStatus === "MERAH"
                                                  ? "inline-block"
                                                  : "none",
                                            }}
                                          ></span>
                                        </div>
                                      </div>

                                      <div className="flex flex-wrap items-center gap-2 mb-2 text-left">
                                        <span
                                          className={`text-[9px] px-2 py-0.5 rounded-full border uppercase tracking-widest ${badgeClass}`}
                                        >
                                          {statusLabel}
                                        </span>
                                        {isTeachingToday && (
                                          <span className="text-[9px] font-black text-zinc-500 bg-zinc-100/80 px-2 py-0.5 border border-zinc-200 rounded-full uppercase tracking-wider">
                                            {teacherSchedulesToday.length} Kelas
                                          </span>
                                        )}
                                      </div>
                                    </div>

                                    <div className="mt-2 pt-3 border-t border-gray-100/60 text-left">
                                      <div className="flex justify-between items-center text-[10px] text-zinc-400 font-extrabold uppercase tracking-widest">
                                        <span>Lihat Detail KBM</span>
                                        <span>
                                          {expandedKamadTeacherDetails[
                                            teacher.nip
                                          ]
                                            ? "▲"
                                            : "▼"}
                                        </span>
                                      </div>

                                      {expandedKamadTeacherDetails[
                                        teacher.nip
                                      ] && (
                                        <div className="mt-3 space-y-2.5 bg-white p-3.5 rounded-2xl border border-gray-150 shadow-xs animate-slideIn">
                                          {!isTeachingToday ? (
                                            <p className="text-[10px] text-zinc-400 font-bold italic">
                                              Guru ini tidak memiliki jadwal
                                              mengajar pada hari {todayDayName}.
                                            </p>
                                          ) : (
                                            teachDetailList.map((dt, sIdx) => {
                                              const isInactive =
                                                dt.state === "INACTIVE";
                                              return (
                                                <div
                                                  key={sIdx}
                                                  className={`border-b border-zinc-50 last:border-b-0 pb-3 last:pb-0 ${isInactive || currentHoliday ? "opacity-70" : ""}`}
                                                >
                                                  <p className="text-[9px] font-black text-zinc-400 uppercase tracking-widest leading-none block text-left">
                                                    {dt.mapel}{" "}
                                                    {currentHoliday ? (
                                                      <span className="text-rose-600 font-bold">
                                                        (HARI LIBUR)
                                                      </span>
                                                    ) : isInactive ? (
                                                      <span className="text-orange-600 font-bold">
                                                        (NONAKTIF / KEGIATAN)
                                                      </span>
                                                    ) : null}
                                                  </p>

                                                  {currentHoliday ? (
                                                    <p className="text-xs font-black text-zinc-500 mt-1 text-left">
                                                      Hari libur{" "}
                                                      <span className="text-rose-700 bg-rose-50 border border-rose-100 px-1.5 py-0.5 rounded-md font-extrabold">
                                                        {currentHoliday.keterangan ||
                                                          "Libur Madrasah"}
                                                      </span>{" "}
                                                      pada JP{" "}
                                                      <span className="font-mono text-zinc-950 bg-zinc-100 px-1 py-0.5 rounded text-[10px]">
                                                        {dt.jps.join(", ")}
                                                      </span>{" "}
                                                      ({dt.startTime} -{" "}
                                                      {dt.endTime})
                                                    </p>
                                                  ) : isInactive ? (
                                                    <p className="text-xs font-black text-zinc-500 mt-1 text-left">
                                                      Mengikuti kegiatan{" "}
                                                      <span className="text-orange-700 bg-orange-50 border border-orange-100 px-1.5 py-0.5 rounded-md font-extrabold">
                                                        {dt.reasonInactive ||
                                                          "Materi Nonaktif"}
                                                      </span>{" "}
                                                      pada JP{" "}
                                                      <span className="font-mono text-zinc-950 bg-zinc-100 px-1 py-0.5 rounded text-[10px]">
                                                        {dt.jps.join(", ")}
                                                      </span>{" "}
                                                      ({dt.startTime} -{" "}
                                                      {dt.endTime})
                                                    </p>
                                                  ) : (
                                                    <p className="text-xs font-black text-zinc-800 mt-1 text-left">
                                                      Mengajar di Kelas{" "}
                                                      <span className="text-green-700">
                                                        {dt.kelas}
                                                      </span>
                                                      , JP{" "}
                                                      <span className="font-mono text-zinc-950 bg-zinc-100 px-1 py-0.5 rounded text-[10px]">
                                                        {dt.jps.join(", ")}
                                                      </span>{" "}
                                                      ({dt.startTime} -{" "}
                                                      {dt.endTime})
                                                    </p>
                                                  )}

                                                  <div className="mt-2 flex items-center gap-2 bg-zinc-50/50 p-2 rounded-xl border border-zinc-100">
                                                    <div
                                                      className={`w-2 h-2 rounded-full shrink-0 ${
                                                        isInactive
                                                          ? "bg-zinc-400 border border-zinc-300"
                                                          : dt.scanned
                                                            ? "bg-green-500"
                                                            : dt.state ===
                                                                "ACTIVE"
                                                              ? "bg-rose-500 animate-pulse"
                                                              : dt.state ===
                                                                  "STANDBY"
                                                                ? "bg-amber-400"
                                                                : "bg-zinc-400"
                                                      }`}
                                                    />
                                                    <p className="text-[10px] font-bold text-zinc-600">
                                                      {isInactive ? (
                                                        <span className="text-zinc-500 font-black">
                                                          Mengikuti Kegiatan:{" "}
                                                          {dt.reasonInactive ||
                                                            "Materi Nonaktif"}
                                                        </span>
                                                      ) : dt.scanned ? (
                                                        <span className="text-green-600 font-black">
                                                          Presensi Jam KBM
                                                          Terdaftar (
                                                          {dt.scanTime} WITA)
                                                        </span>
                                                      ) : dt.state ===
                                                        "ACTIVE" ? (
                                                        <span className="text-rose-600 font-black animate-pulse">
                                                          Belum Melakukan Scan
                                                          di Jam Pelajaran Ini
                                                        </span>
                                                      ) : dt.state ===
                                                        "STANDBY" ? (
                                                        <span className="text-amber-600 font-black">
                                                          Belum Scan (Standby -
                                                          Kelas Belum Dimulai)
                                                        </span>
                                                      ) : (
                                                        <span className="text-zinc-500 font-black">
                                                          Tidak Melakukan Scan
                                                          (Melewati Batas JP)
                                                        </span>
                                                      )}
                                                    </p>
                                                  </div>
                                                </div>
                                              );
                                            })
                                          )}
                                        </div>
                                      )}
                                    </div>
                                  </div>
                                );
                              })}
                            </div>
                          </div>
                        );
                      })()}
                    </>
                  ) : (
                    // Regular Teacher Dashboard View: Agenda Mengajar Hari Ini
                    <>
                      <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 border-b border-gray-50 pb-5">
                        <div>
                          <span className="text-[9px] font-black text-green-700 uppercase tracking-widest bg-green-50 px-2.5 py-1 rounded-full border border-green-100">
                            Jadwal Mengajar Real-time
                          </span>
                          <h3 className="text-xl font-black text-zinc-900 mt-2">
                            Agenda Mengajar Hari Ini
                          </h3>
                          <p className="text-xs text-gray-400 font-bold mt-1">
                            Sistem menyala otomatis sesuai dengan jam pelajaran
                            yang ditetapkan oleh Admin.
                          </p>
                        </div>
                        {/* Live Clock indicator representing why it is active real-time */}
                        <div className="flex items-center gap-3 bg-zinc-950 text-white px-5 py-3 rounded-2xl shadow-md shrink-0">
                          <Clock
                            size={16}
                            className="text-green-400 animate-pulse"
                          />
                          <div className="flex flex-col text-left">
                            <span className="text-[8px] font-black text-gray-400 uppercase tracking-widest">
                              Waktu Presensi
                            </span>
                            <span className="text-sm font-black font-mono tracking-wider">
                              {currentTime || "00:00"} WITA
                            </span>
                          </div>
                        </div>
                      </div>

                      {(() => {
                        const IndonesianDays = [
                          "Minggu",
                          "Senin",
                          "Selasa",
                          "Rabu",
                          "Kamis",
                          "Jumat",
                          "Sabtu",
                        ];
                        const todayDayName =
                          IndonesianDays[new Date().getDay()];

                        if (todayDayName === "Minggu") {
                          return (
                            <div className="py-8 text-center text-gray-400 italic font-medium">
                              Hari Minggu tidak ada jadwal KBM. Selamat
                              berlibur!
                            </div>
                          );
                        }

                        // Get currently logged-in teacher's schedules for today
                        const todayDateStr = new Date().toLocaleDateString(
                          "en-CA",
                        );
                        const myTodaySchedules = teachingSchedules.filter(
                          (s) =>
                            s.nip === session?.uid && s.hari === todayDayName,
                        );

                        if (myTodaySchedules.length === 0) {
                          return (
                            <div className="py-8 text-center text-gray-400 italic font-bold uppercase tracking-wider text-xs">
                              Anda tidak memiliki jadwal mengajar terdaftar pada
                              hari {todayDayName}.
                            </div>
                          );
                        }

                        // Get today's daily setting
                        const todaySetting = settings.find(
                          (st) => st.hari === todayDayName,
                        );
                        const currentHoliday = holidays.find(
                          (h) => h.tanggal === todayDateStr,
                        );

                        // Sort schedule sequentially based on minimum JP assigned
                        const sortedSchedules = [...myTodaySchedules].sort(
                          (a, b) => {
                            const minA =
                              a.jps && a.jps.length > 0
                                ? Math.min(...a.jps)
                                : 99;
                            const minB =
                              b.jps && b.jps.length > 0
                                ? Math.min(...b.jps)
                                : 99;
                            return minA - minB;
                          },
                        );

                        return (
                          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                            {sortedSchedules.map((sched) => {
                              const schedJps = sched.jps || [];
                              const minJp =
                                schedJps.length > 0
                                  ? Math.min(...schedJps)
                                  : null;
                              const maxJp =
                                schedJps.length > 0
                                  ? Math.max(...schedJps)
                                  : null;

                              let startTime = "";
                              let endTime = "";

                              if (minJp && maxJp && todaySetting?.jpTimes) {
                                startTime =
                                  todaySetting.jpTimes[minJp]?.start || "";
                                endTime =
                                  todaySetting.jpTimes[maxJp]?.end || "";
                              }

                              // Check if teacher has already marked presence for this today
                              const attendanceRecord = teacherAttendance.find(
                                (ta) =>
                                  ta.nip === session?.uid &&
                                  ta.kelas === sched.kelas &&
                                  ta.tanggal === todayDateStr,
                              );

                              // Check if there is ANY manual leave (Izin/Sakit/Alfa) record for the teacher today, either general or classy
                              const leaveRecord = teacherAttendance.find(
                                (ta) =>
                                  ta.nip === session?.uid &&
                                  ta.tanggal === todayDateStr &&
                                  (ta.status === "Izin" ||
                                    ta.status === "Sakit" ||
                                    ta.status === "Alfa"),
                              );

                              const isLeaveOrPaidOff = !!leaveRecord;
                              const leaveStatus = leaveRecord?.status;
                              const leaveKeterangan =
                                leaveRecord?.keterangan || "";

                              const isAlreadyAttended =
                                !!attendanceRecord && !leaveRecord;

                              // Check if any of these JPs are inactive in today's settings
                              const activeJpsList =
                                todaySetting?.activeJps &&
                                Array.isArray(todaySetting.activeJps)
                                  ? todaySetting.activeJps
                                  : Array.from(
                                      {
                                        length:
                                          todayDayName === "Jumat" ? 6 : 8,
                                      },
                                      (_, i) => i + 1,
                                    );
                              const isJpInactive = schedJps.some(
                                (jp) => !activeJpsList.includes(jp),
                              );

                              // Evaluate dynamic button availability
                              let buttonStatus:
                                | "not_started"
                                | "active"
                                | "ended"
                                | "inactive" = "active";
                              let statusLabel = "Siap Mulai Presensi";
                              let statusBg =
                                "bg-amber-50 border-amber-200 text-amber-700";

                              if (currentHoliday) {
                                buttonStatus = "inactive";
                                statusLabel = `Hari Libur: ${currentHoliday.keterangan}`;
                                statusBg =
                                  "bg-rose-50 border-rose-200 text-rose-700 font-extrabold";
                              } else if (isJpInactive) {
                                buttonStatus = "inactive";
                                statusLabel = todaySetting?.reasonInactive
                                  ? `${todaySetting.reasonInactive}`
                                  : "Jam Pelajaran Nonaktif";
                                statusBg =
                                  "bg-zinc-100 border-zinc-250 text-zinc-500 font-extrabold";
                              } else if (startTime && endTime && currentTime) {
                                const currentMin =
                                  parseInt(currentTime.split(":")[0]) * 60 +
                                  parseInt(currentTime.split(":")[1]);
                                const startMin =
                                  parseInt(startTime.split(":")[0]) * 60 +
                                  parseInt(startTime.split(":")[1]);
                                const endMin =
                                  parseInt(endTime.split(":")[0]) * 60 +
                                  parseInt(endTime.split(":")[1]);

                                if (currentMin < startMin) {
                                  buttonStatus = "not_started";
                                  statusLabel = "Belum Waktunya";
                                  statusBg =
                                    "bg-zinc-100 border-zinc-200 text-zinc-500";
                                } else if (currentMin > endMin) {
                                  buttonStatus = "ended";
                                  statusLabel = "Sesi Selesai";
                                  statusBg =
                                    "bg-rose-50 border-rose-150 text-rose-600";
                                } else {
                                  buttonStatus = "active";
                                  statusLabel = "Jam Mengajar Aktif";
                                  statusBg =
                                    "bg-green-50 border-green-200 text-green-700 font-extrabold animate-pulse";
                                }
                              } else {
                                statusLabel =
                                  "Waktu JP Belum Diatur (Standar Aktif)";
                                statusBg =
                                  "bg-blue-50 border-blue-150 text-blue-600";
                              }

                              return (
                                <div
                                  key={sched.id}
                                  className={`p-6 rounded-3xl border transition-all flex flex-col justify-between h-full relative overflow-hidden group ${
                                    currentHoliday
                                      ? "bg-rose-50/10 border-rose-150 opacity-80"
                                      : isJpInactive
                                        ? "bg-zinc-105 border-zinc-200/80 opacity-65"
                                        : isLeaveOrPaidOff
                                          ? leaveStatus === "Izin"
                                            ? "bg-sky-50/25 border-sky-150 shadow-sm"
                                            : leaveStatus === "Sakit"
                                              ? "bg-amber-50/25 border-amber-150 shadow-sm"
                                              : "bg-rose-50/25 border-rose-150 shadow-sm"
                                          : isAlreadyAttended
                                            ? "bg-emerald-50/25 border-emerald-100 hover:bg-emerald-50/40"
                                            : buttonStatus === "active"
                                              ? "bg-white border-green-300 ring-2 ring-green-100 shadow-md transform -translate-y-1"
                                              : "bg-zinc-50/45 border-zinc-100 opacity-80 hover:opacity-100"
                                  }`}
                                >
                                  <div>
                                    <div className="flex justify-between items-start mb-4 gap-2">
                                      <div className="flex flex-col gap-1 text-left">
                                        <span className="text-[10px] font-black text-gray-400 uppercase tracking-widest">
                                          {sched.mapel || "Mata Pelajaran"}{" "}
                                          {currentHoliday ? (
                                            <span className="text-rose-600 font-extrabold border border-rose-100 bg-rose-50/50 px-1 py-0.5 rounded text-[8px]">
                                              (HARI LIBUR)
                                            </span>
                                          ) : isJpInactive ? (
                                            <span className="text-orange-600 font-extrabold border border-orange-100 bg-orange-50/50 px-1 py-0.5 rounded text-[8px]">
                                              (KEGIATAN)
                                            </span>
                                          ) : null}
                                        </span>
                                        <span className="text-lg font-black text-zinc-900 group-hover:text-green-950 transition-colors">
                                          Kelas {sched.kelas}
                                        </span>
                                      </div>
                                      <div
                                        className={`px-2.5 py-1 rounded-full text-[9px] font-black uppercase tracking-wider border whitespace-nowrap leading-none ${
                                          isLeaveOrPaidOff
                                            ? leaveStatus === "Izin"
                                              ? "bg-sky-50 border-sky-250 text-sky-700"
                                              : leaveStatus === "Sakit"
                                                ? "bg-amber-50 border-amber-250 text-amber-700"
                                                : "bg-rose-50 border-rose-250 text-rose-700"
                                            : isAlreadyAttended
                                              ? "bg-emerald-50 border-emerald-250 text-emerald-700"
                                              : statusBg
                                        }`}
                                      >
                                        {isLeaveOrPaidOff
                                          ? `STATUS: ${leaveStatus?.toUpperCase()}`
                                          : isAlreadyAttended
                                            ? "Selesai Presensi ✓"
                                            : statusLabel}
                                      </div>
                                    </div>

                                    <div className="flex flex-wrap items-center gap-2 mb-4">
                                      <div className="bg-zinc-100 px-2.5 py-1 rounded-xl text-[10px] font-black text-zinc-600 uppercase tracking-wide">
                                        JP {schedJps.join(", ")}
                                      </div>
                                      {startTime && endTime ? (
                                        <div className="bg-zinc-100 px-2.5 py-1 rounded-xl text-[10px] font-black text-zinc-600 font-mono tracking-wide flex items-center gap-1">
                                          <Clock
                                            size={10}
                                            className="text-zinc-500"
                                          />
                                          {startTime} - {endTime}
                                        </div>
                                      ) : (
                                        <span className="text-[8px] text-zinc-400 uppercase font-black tracking-widest leading-none">
                                          Standard JP
                                        </span>
                                      )}
                                    </div>
                                  </div>

                                  <div className="mt-4 pt-4 border-t border-gray-100/60 font-sans">
                                    {isLeaveOrPaidOff ? (
                                      <div
                                        className={`p-4 rounded-2xl border transition-all text-left ${
                                          leaveStatus === "Izin"
                                            ? "bg-sky-50/65 border-sky-150 text-sky-950"
                                            : leaveStatus === "Sakit"
                                              ? "bg-amber-50/65 border-amber-150 text-amber-950"
                                              : "bg-rose-50/65 border-rose-150 text-rose-950"
                                        }`}
                                      >
                                        <div className="flex items-start gap-3">
                                          <div
                                            className={`w-8 h-8 rounded-xl shrink-0 font-bold flex items-center justify-center text-white ${
                                              leaveStatus === "Izin"
                                                ? "bg-sky-600"
                                                : leaveStatus === "Sakit"
                                                  ? "bg-amber-600"
                                                  : "bg-rose-600"
                                            }`}
                                          >
                                            {leaveStatus === "Izin"
                                              ? "✉️"
                                              : leaveStatus === "Sakit"
                                                ? "🤒"
                                                : "🚫"}
                                          </div>
                                          <div className="text-left font-sans flex-1">
                                            <p className="text-xs font-black uppercase">
                                              Berhalangan: {leaveStatus}
                                            </p>
                                            <p className="text-[10px] font-bold text-gray-500 mt-1 leading-normal">
                                              Anda berstatus{" "}
                                              {leaveStatus?.toLowerCase()} hari
                                              ini{" "}
                                              {leaveKeterangan
                                                ? `(${leaveKeterangan})`
                                                : ""}
                                              . Sistem menonaktifkan presensi
                                              kelas karena izin yang diajukan
                                              telah disetujui Kamad.
                                            </p>
                                          </div>
                                        </div>
                                      </div>
                                    ) : isAlreadyAttended ? (
                                      <div className="flex items-center justify-between bg-emerald-50 border border-emerald-200 rounded-2xl p-3.5">
                                        <div className="flex items-center gap-2.5 text-left">
                                          <div className="w-7 h-7 rounded-lg bg-emerald-500 flex items-center justify-center text-white shrink-0 font-extrabold text-sm">
                                            ✓
                                          </div>
                                          <div>
                                            <p className="text-xs font-black text-emerald-950 uppercase leading-none">
                                              Sudah Presensi
                                            </p>
                                            <p className="text-[9px] font-bold text-emerald-600 uppercase tracking-widest mt-1">
                                              Jam check-in:{" "}
                                              {attendanceRecord?.jam || ""}
                                            </p>
                                          </div>
                                        </div>
                                        <span className="text-[8px] font-black text-emerald-500 uppercase tracking-widest bg-white border border-emerald-150 px-2 py-0.5 rounded-full shadow-xs">
                                          Sesi Valid
                                        </span>
                                      </div>
                                    ) : currentHoliday ? (
                                      <div className="bg-rose-50 border border-rose-100/60 rounded-2xl p-3.5 flex items-center gap-3 w-full text-left">
                                        <div className="w-7 h-7 rounded-lg bg-rose-500 flex items-center justify-center text-white shrink-0 font-bold text-[10px] leading-none">
                                          🏖️
                                        </div>
                                        <div className="text-left font-sans">
                                          <p className="text-xs font-black text-rose-950 uppercase leading-none">
                                            Hari Libur Madrasah
                                          </p>
                                          <p className="text-[10px] font-bold text-rose-700 uppercase tracking-wider mt-1 leading-normal">
                                            Hari Libur:{" "}
                                            {currentHoliday.keterangan ||
                                              "Libur Madrasah"}
                                          </p>
                                        </div>
                                      </div>
                                    ) : isJpInactive ? (
                                      <div className="bg-orange-50 border border-orange-100/60 rounded-2xl p-3.5 flex items-center gap-3">
                                        <div className="w-7 h-7 rounded-lg bg-orange-550 flex items-center justify-center text-white shrink-0 font-bold text-xs">
                                          ℹ
                                        </div>
                                        <div className="text-left font-sans">
                                          <p className="text-xs font-black text-orange-950 uppercase leading-none">
                                            Mengikuti Kegiatan
                                          </p>
                                          <p className="text-[9px] font-bold text-orange-600 uppercase tracking-widest mt-1 line-clamp-1">
                                            {todaySetting?.reasonInactive ||
                                              "Materi JP Nonaktif"}
                                          </p>
                                        </div>
                                      </div>
                                    ) : (
                                      <div
                                        className={`p-4 rounded-2xl border transition-all ${
                                          buttonStatus === "active"
                                            ? "bg-amber-50/50 border-amber-200/80 text-amber-900"
                                            : "bg-zinc-50 border-zinc-200 text-zinc-500"
                                        }`}
                                      >
                                        <div className="flex items-start gap-3">
                                          <div
                                            className={`w-8 h-8 rounded-xl shrink-0 font-bold flex items-center justify-center ${
                                              buttonStatus === "active"
                                                ? "bg-amber-500 text-white animate-pulse"
                                                : "bg-zinc-200 text-zinc-500"
                                            }`}
                                          >
                                            <QrCode size={16} />
                                          </div>
                                          <div className="text-left font-sans flex-1">
                                            <p
                                              className={`text-xs font-black uppercase ${
                                                buttonStatus === "active"
                                                  ? "text-amber-950"
                                                  : "text-zinc-700"
                                              }`}
                                            >
                                              Wajib Scanner Presensi
                                            </p>
                                            <p className="text-[10px] font-bold text-gray-400 mt-1 leading-normal">
                                              {buttonStatus === "not_started"
                                                ? "Belum masuk jam pelajaran. Silakan scan kartu/QR Code pada alat scanner saat jam pelajaran dimulai."
                                                : buttonStatus === "active"
                                                  ? "Lakukan presensi dengan memindai kartu presensi atau QR Code Anda langsung pada mesin scanner kelas."
                                                  : "Waktu mengajar Anda untuk sesi kelas ini telah berakhir."}
                                            </p>
                                          </div>
                                        </div>
                                      </div>
                                    )}
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        );
                      })()}
                    </>
                  )}
                </div>
              )}

              <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
                <div className="lg:col-span-2 bg-white p-8 rounded-[2rem] shadow-sm border border-gray-100">
                  <div className="flex justify-between items-center mb-6">
                    <div>
                      <h3 className="text-sm font-black text-green-900 uppercase tracking-widest">
                        Grafik Kehadiran
                      </h3>
                      <p className="text-[10px] text-gray-400 font-bold">
                        Statistik kehadiran 7 hari terakhir
                      </p>
                    </div>
                    <BarChart3 className="text-green-600" size={20} />
                  </div>
                  {stats.chartData ? (
                    <AttendanceChart data={stats.chartData} />
                  ) : (
                    <div className="h-64 flex items-center justify-center text-gray-300 italic text-xs">
                      Memuat data grafik...
                    </div>
                  )}
                </div>

                <div className="bg-green-950 p-8 rounded-[2rem] text-white flex flex-col justify-between overflow-hidden relative group">
                  <div className="relative z-10">
                    <h3 className="text-lg font-black leading-tight mb-2">
                      Selamat Datang,
                      <br />
                      {session?.name}!
                    </h3>
                    <p className="text-green-300 text-[10px] font-bold uppercase tracking-widest">
                      Akses Panel {session?.jabatan || session?.role}
                    </p>
                  </div>

                  <div className="mt-8 space-y-4 relative z-10">
                    <div className="bg-white/10 p-4 rounded-2xl backdrop-blur-sm border border-white/10">
                      <p className="text-[10px] font-black text-green-200 uppercase tracking-widest mb-1">
                        Status Sistem
                      </p>
                      <div className="flex items-center gap-2">
                        <div className="w-2 h-2 rounded-full bg-green-400 animate-pulse"></div>
                        <span className="text-xs font-bold">
                          Online / Real-time
                        </span>
                      </div>
                    </div>
                  </div>

                  <School
                    size={120}
                    className="absolute -bottom-6 -right-6 text-white/5 group-hover:scale-110 transition-transform duration-500"
                  />
                </div>
              </div>
            </motion.div>
          )}

          {activePanel === "scanner" && (
            <motion.div
              key="scanner"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              className="flex flex-col items-center"
            >
              <div className="text-center mb-6">
                <div className="inline-block px-4 py-2 bg-white rounded-full shadow-sm border border-zinc-100 mb-2">
                  <span className="text-xs font-black text-green-900 uppercase tracking-widest flex items-center gap-2">
                    <Calendar size={14} />
                    {new Date().toLocaleDateString("id-ID", {
                      weekday: "long",
                      day: "numeric",
                      month: "long",
                      year: "numeric",
                    })}
                  </span>
                </div>
                <h2 className="text-2xl font-black text-green-950">
                  Presensi Digital
                </h2>
                <p className="text-zinc-500 text-sm font-medium">
                  Arahkan barcode kartu ke kamera
                </p>
              </div>

              <div className="bg-white p-6 rounded-3xl shadow-xl w-full max-w-lg">
                <div className="flex gap-2 p-1 bg-gray-100 rounded-xl mb-4">
                  <button
                    onClick={() => {
                      setScanType("Siswa");
                      setScanResult(null);
                      setScanning(false);
                      setClassScanFailCount(0);
                    }}
                    className={`flex-1 py-2 text-sm font-bold rounded-lg transition-all ${scanType === "Siswa" ? "bg-green-800 text-white shadow" : "text-gray-500"}`}
                  >
                    Scan Siswa
                  </button>
                  <button
                    onClick={() => {
                      setScanType("Kelas");
                      setScanResult(null);
                      setScanning(false);
                      setClassScanFailCount(0);
                    }}
                    className={`flex-1 py-2 text-sm font-bold rounded-lg transition-all ${scanType === "Kelas" ? "bg-blue-800 text-white shadow" : "text-gray-500"}`}
                  >
                    Scan Kelas
                  </button>
                </div>

                <div className="aspect-square bg-zinc-900 rounded-2xl flex flex-col items-center justify-center relative overflow-hidden mb-4 group ring-8 ring-gray-50">
                  {scanning ? (
                    <div id="reader" className="w-full h-full"></div>
                  ) : (
                    <>
                      <QrCode
                        size={120}
                        className="text-white/20 group-hover:scale-110 transition-transform duration-500"
                      />
                      <div className="absolute inset-0 border-2 border-dashed border-green-500/50 m-8 rounded-2xl animate-pulse"></div>
                      <div className="absolute top-4 left-1/2 -translate-x-1/2 bg-white/10 backdrop-blur-md px-3 py-1 rounded-full text-[10px] text-white font-bold uppercase tracking-widest border border-white/10">
                        Mode:{" "}
                        {scanType === "Siswa"
                          ? "Absensi Siswa"
                          : "Absensi Guru (Mengajar)"}
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

                {scanType !== "Kelas" || classScanFailCount >= 2 ? (
                  <div className="mb-4">
                    <div className="flex items-center justify-between mb-1 ml-1">
                      <div className="flex items-center gap-1.5">
                        <label className="block text-[10px] font-black text-gray-400 uppercase">
                          Input Manual
                        </label>
                        {scanType === "Kelas" && (
                          <span className="bg-amber-100 text-amber-800 font-extrabold text-[9px] px-2 py-0.5 rounded-full uppercase tracking-wider">
                            Terbuka (Scan Gagal {classScanFailCount}x)
                          </span>
                        )}
                      </div>
                      {scanning && (
                        <button
                          onClick={() => setScanning(false)}
                          className="text-[10px] font-black text-red-500 uppercase"
                        >
                          Matikan Kamera
                        </button>
                      )}
                    </div>
                    <div className="flex gap-2">
                      <input
                        type="text"
                        value={scanInput}
                        onChange={(e) => setScanInput(e.target.value)}
                        placeholder={
                          scanType === "Siswa"
                            ? "Contoh NISN: 12345"
                            : "Masukkan Nama Kelas"
                        }
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
                ) : (
                  <div className="mb-4 p-5 bg-zinc-50 border border-dashed border-zinc-200 rounded-2xl text-center relative overflow-hidden">
                    <div className="absolute top-2 right-2">
                      {scanning && (
                        <button
                          onClick={() => setScanning(false)}
                          className="text-[9px] font-black text-red-500 hover:text-red-700 uppercase tracking-wider bg-white px-2 py-1 rounded-md border border-zinc-150"
                        >
                          Matikan Kamera
                        </button>
                      )}
                    </div>
                    <span className="text-[10px] uppercase font-black text-zinc-400 tracking-widest block mb-1">
                      Input Manual Kelas Terkunci
                    </span>
                    <p className="text-[11px] text-zinc-500 font-medium">
                      Harap lakukan scan QR Code Kelas terlebih dahulu
                      menggunakan kamera.
                    </p>
                    <p className="text-[9px] text-zinc-400 mt-1.5 font-bold">
                      Scan gagal kamera saat ini:{" "}
                      <span className="text-zinc-600 font-black">
                        {classScanFailCount}/2
                      </span>
                    </p>
                    <p className="text-[9px] text-zinc-350 mt-1 italic leading-relaxed">
                      *Input manual otomatis terbuka jika scan kamera gagal
                      sebanyak 2 kali untuk membuktikan guru berada di dalam
                      kelas.
                    </p>
                  </div>
                )}

                {scanResult && (
                  <motion.div
                    initial={{ scale: 0.95, opacity: 0 }}
                    animate={{ scale: 1, opacity: 1 }}
                    className={`p-4 rounded-2xl border-2 flex flex-col gap-2`}
                    style={{
                      ...(scanResult.status === "Alfa"
                        ? {
                            backgroundColor: "#fff1f2",
                            borderColor: "#fecdd3",
                            color: "#9f1239",
                          }
                        : scanResult.success
                          ? {
                              backgroundColor: "#f0fdf4",
                              borderColor: "#bbf7d0",
                              color: "#166534",
                            }
                          : {
                              backgroundColor: "#fef2f2",
                              borderColor: "#fecaca",
                              color: "#991b1b",
                            }),
                    }}
                  >
                    <div className="flex items-center gap-3">
                      <div
                        className={`w-8 h-8 rounded-full flex items-center justify-center ${scanResult.status === "Alfa" ? "bg-rose-500" : scanResult.success ? "bg-green-500" : "bg-red-500"} text-white`}
                      >
                        {scanResult.success ? (
                          <CheckCircle2 size={16} />
                        ) : (
                          <AlertTriangle size={16} />
                        )}
                      </div>
                      <div className="flex-1">
                        <p className="font-bold">
                          {scanResult.status === "Alfa"
                            ? "Gagal: Sudah Jam Pulang"
                            : scanResult.success
                              ? "Berhasil Terdaftar"
                              : "Gagal Memproses"}
                        </p>
                        <p className="text-xs opacity-75">
                          {scanResult.message}
                        </p>
                      </div>
                    </div>
                    {scanResult.success &&
                      scanResult.message.includes("Terlambat") && (
                        <div className="bg-orange-500 text-white p-3 rounded-xl flex items-center gap-3 animate-bounce">
                          <Clock size={20} className="animate-spin-slow" />
                          <div>
                            <p className="text-xs font-black uppercase tracking-widest">
                              Peringatan Terlambat!
                            </p>
                            <p className="text-[10px] font-bold">
                              Waktu Anda akan tercatat dalam akumulasi
                              keterlambatan.
                            </p>
                          </div>
                        </div>
                      )}
                  </motion.div>
                )}
              </div>
              <p className="mt-4 text-xs text-zinc-400 font-medium">
                Pastikan QR Code berada di tengah kotak pemindai.
              </p>
            </motion.div>
          )}

          {activePanel === "siswa" && (
            <motion.div
              key="siswa"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
            >
              <div className="mb-4 flex flex-wrap gap-3 justify-between items-center">
                <div className="flex flex-wrap gap-3 items-center">
                  <div className="relative w-64">
                    <Search
                      size={16}
                      className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400"
                    />
                    <input
                      type="text"
                      placeholder="Cari Siswa..."
                      value={searchTermSiswa}
                      onChange={(e) => {
                        setSearchTermSiswa(e.target.value);
                        setPagination({ ...pagination, siswa: 0 });
                      }}
                      className="w-full bg-white border border-gray-100 rounded-xl py-2 pl-10 pr-4 text-sm focus:ring-2 focus:ring-green-500"
                    />
                  </div>
                  <select
                    value={filterAdminSiswaClass}
                    onChange={(e) => setFilterAdminSiswaClass(e.target.value)}
                    className="bg-white border border-gray-100 rounded-xl px-4 py-2 text-sm font-bold text-gray-700"
                  >
                    <option value="">Semua Kelas</option>
                    {classrooms.map((c) => (
                      <option key={c.nama} value={c.nama}>
                        {c.nama}
                      </option>
                    ))}
                  </select>
                </div>
                <div className="flex gap-2">
                  <select
                    value={pageSize}
                    onChange={(e) => setPageSize(parseInt(e.target.value))}
                    className="bg-zinc-100 text-zinc-600 px-3 py-2 rounded-xl text-[10px] font-black uppercase border-0 focus:ring-0"
                  >
                    {[10, 20, 50, 100].map((v) => (
                      <option key={v} value={v}>
                        Tampil {v}
                      </option>
                    ))}
                  </select>
                  <button
                    onClick={() => downloadTemplate("Siswa")}
                    className="bg-zinc-100 text-zinc-600 px-4 py-2 rounded-xl text-[10px] font-black tracking-widest uppercase flex items-center gap-2 hover:bg-zinc-200 transition-all"
                  >
                    <Download size={14} /> Template
                  </button>
                  <button
                    onClick={() => handleImportExcel("Siswa")}
                    className="bg-blue-100 text-blue-600 px-4 py-2 rounded-xl text-[10px] font-black tracking-widest uppercase flex items-center gap-2 hover:bg-blue-200 transition-all"
                  >
                    <Upload size={14} /> Import
                  </button>
                  <button
                    onClick={downloadMassSiswaKartu}
                    className="bg-purple-800 text-white px-4 py-2 rounded-xl text-[10px] font-black tracking-widest uppercase flex items-center gap-2 hover:bg-purple-700 transition-all shadow-sm"
                  >
                    <Download size={14} /> Unduh Kartu
                  </button>
                  <button
                    onClick={() => setShowNaikKelasModal(true)}
                    className="bg-orange-600 text-white px-4 py-2 rounded-xl text-[10px] font-black tracking-widest uppercase flex items-center gap-2 hover:bg-orange-500 transition-all shadow-sm"
                  >
                    <ArrowUpCircle size={14} /> Naik Kelas
                  </button>
                  <button
                    onClick={() => {
                      setEditingSiswa({
                        nisn: "",
                        nama: "",
                        jenisKelamin: "L",
                        tempat: "",
                        tgl: "",
                        kelas: "",
                        ayah: "",
                        ibu: "",
                        hp: "",
                      });
                      setShowSiswaModal(true);
                    }}
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
                      {filteredSiswa
                        .filter((s) =>
                          filterAdminSiswaClass
                            ? s.kelas === filterAdminSiswaClass
                            : true,
                        )
                        .slice(pagination.siswa, pagination.siswa + pageSize)
                        .map((s, i) => (
                          <tr
                            key={s.nisn}
                            className="hover:bg-gray-50 transition-colors group"
                          >
                            <td className="px-6 py-4 text-center font-bold text-gray-400 text-xs">
                              {pagination.siswa + i + 1}
                            </td>
                            <td className="px-6 py-4">
                              {s.foto ? (
                                <img
                                  src={s.foto}
                                  className="w-8 h-8 rounded-lg object-cover ring-2 ring-gray-100"
                                  alt={s.nama}
                                />
                              ) : (
                                <div className="w-8 h-8 rounded-lg bg-gray-100 flex items-center justify-center text-gray-400">
                                  <User size={14} />
                                </div>
                              )}
                            </td>
                            <td className="px-6 py-4 font-mono font-bold text-gray-400">
                              {s.nisn}
                            </td>
                            <td className="px-6 py-4 font-bold">{s.nama}</td>
                            <td className="px-6 py-4">
                              <span className="text-[10px] font-bold">
                                {s.kelas}
                              </span>
                            </td>
                            <td className="px-6 py-4">
                              <p className="text-[10px] font-medium text-gray-500">
                                A: {s.ayah || "-"}
                              </p>
                              <p className="text-[10px] font-medium text-gray-500">
                                I: {s.ibu || "-"}
                              </p>
                            </td>
                            <td className="px-6 py-4">{s.hp}</td>
                            <td className="px-6 py-4">
                              <div className="flex gap-2">
                                <button
                                  onClick={() => setSelectedStudentCard(s)}
                                  className="text-green-600 hover:text-green-800"
                                >
                                  <QrCode size={16} />
                                </button>
                                <button
                                  onClick={() => {
                                    setEditingSiswa(s);
                                    setShowSiswaModal(true);
                                  }}
                                  className="text-blue-500 hover:text-blue-800"
                                >
                                  <Edit size={16} />
                                </button>
                                <button
                                  onClick={() =>
                                    handleDeleteSiswa(s.nisn, s.nama)
                                  }
                                  className="text-red-500 hover:text-red-800"
                                >
                                  <Trash2 size={16} />
                                </button>
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
                  <span className="text-[10px] font-black text-gray-400 uppercase tracking-widest">
                    Halaman {Math.floor(pagination.siswa / pageSize) + 1}
                  </span>
                  <div className="flex gap-2">
                    <button
                      disabled={pagination.siswa === 0}
                      onClick={() =>
                        setPagination({
                          ...pagination,
                          siswa: pagination.siswa - pageSize,
                        })
                      }
                      className="p-2 rounded-lg bg-white border border-gray-100 text-gray-400 disabled:opacity-30 hover:bg-gray-50 transition-all"
                    >
                      <ChevronLeft size={16} />
                    </button>
                    <button
                      disabled={
                        pagination.siswa + pageSize >= filteredSiswa.length
                      }
                      onClick={() =>
                        setPagination({
                          ...pagination,
                          siswa: pagination.siswa + pageSize,
                        })
                      }
                      className="p-2 rounded-lg bg-white border border-gray-100 text-gray-400 disabled:opacity-30 hover:bg-gray-50 transition-all"
                    >
                      <ChevronRight size={16} />
                    </button>
                  </div>
                </div>
              )}
            </motion.div>
          )}

          {activePanel === "analisis" && (
            <motion.div
              key="analisis"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              className="space-y-8"
            >
              <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 bg-white p-6 rounded-[2rem] shadow-sm border border-zinc-100 mb-8 text-zinc-900">
                <div className="flex flex-col">
                  <h1 className="text-2xl font-black text-green-950">
                    Analisis Kehadiran
                  </h1>
                  <p className="text-zinc-500 font-medium">
                    Monitoring pergerakan data presensi sekolah
                  </p>
                </div>
                <div className="flex items-center gap-3">
                  <div className="flex flex-col">
                    <label className="text-[9px] font-black text-zinc-400 uppercase tracking-widest mb-1 pl-1">
                      Pilih Bulan
                    </label>
                    <input
                      type="month"
                      value={analysisMonth}
                      onChange={(e) => setAnalysisMonth(e.target.value)}
                      className="bg-zinc-100 border-0 rounded-xl px-3 py-2 font-bold text-xs"
                    />
                  </div>
                  <div className="flex flex-col">
                    <label className="text-[9px] font-black text-zinc-400 uppercase tracking-widest mb-1 pl-1">
                      Baris
                    </label>
                    <select
                      value={pageSize}
                      onChange={(e) => setPageSize(parseInt(e.target.value))}
                      className="bg-zinc-100 text-zinc-600 px-3 py-2 rounded-xl text-xs font-bold border-0 focus:ring-0"
                    >
                      {[10, 20, 50, 100].map((v) => (
                        <option key={v} value={v}>
                          {v}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div className="flex flex-col relative justify-end">
                    <label className="text-[9px] font-black text-zinc-400 uppercase tracking-widest mb-1 pl-1">
                      Kelas
                    </label>
                    <div className="relative flex items-center gap-2">
                      {session?.role === "Guru" && session?.isWali && (
                        <span className="absolute -top-3 right-0 bg-red-100 text-red-600 text-[8px] font-black px-2 py-0.5 rounded-full uppercase tracking-widest border border-red-200">
                          Terkunci
                        </span>
                      )}
                      <select
                        value={analysisClass}
                        onChange={(e) => setAnalysisClass(e.target.value)}
                        disabled={session?.role === "Guru" && session?.isWali}
                        className="bg-zinc-50 border-0 rounded-xl px-4 py-2 font-bold text-xs disabled:opacity-50 disabled:cursor-not-allowed min-w-[120px]"
                      >
                        {!(session?.role === "Guru" && session?.isWali) && (
                          <option value="">Semua Kelas</option>
                        )}
                        {(classrooms || [])
                          .filter((c) =>
                            session?.role === "Guru" && session?.isWali
                              ? c.nama === session?.kelas
                              : true,
                          )
                          .map((c) => (
                            <option key={c.nama} value={c.nama}>
                              {c.nama}
                            </option>
                          ))}
                      </select>
                    </div>
                  </div>
                </div>
              </div>
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
                {/* Siswa Attendance Analysis */}
                <div className="bg-white p-6 rounded-3xl shadow-sm border border-gray-100 min-h-[400px]">
                  <h3 className="text-sm font-black text-green-900 uppercase tracking-widest mb-6">
                    {session?.role === "Guru" && session?.isWali
                      ? `Analisis Kehadiran Siswa Kelas ${session.kelas}`
                      : "Analisis Kehadiran Siswa Per Kelas"}
                  </h3>
                  <div className="h-64 relative">
                    {classrooms.length > 0 ? (
                      <ResponsiveContainer
                        width="100%"
                        height="100%"
                        minWidth={0}
                      >
                        <BarChart
                          data={classrooms
                            .filter((c) =>
                              analysisClass ? c.nama === analysisClass : true,
                            )
                            .map((c) => {
                              const currentMonthPrefix = analysisMonth;
                              const classAttendance = (attendance || []).filter(
                                (a) =>
                                  a.kelas === c.nama &&
                                  a.tanggal.startsWith(currentMonthPrefix),
                              );
                              const totalSiswa = (students || []).filter(
                                (s) => s.kelas === c.nama,
                              ).length;

                              const daysPassed =
                                effectiveDaysForAnalysisMonth.length;
                              const totalPeluang = totalSiswa * daysPassed;
                              const hadirCount = classAttendance.filter(
                                (a) =>
                                  a.status === "Hadir" ||
                                  a.status === "Izin" ||
                                  a.status === "Sakit",
                              ).length;
                              return {
                                name: c.nama || "Kelas",
                                percentage:
                                  totalPeluang > 0
                                    ? Math.round(
                                        (hadirCount / totalPeluang) * 100,
                                      )
                                    : 0,
                              };
                            })}
                        >
                          <CartesianGrid
                            strokeDasharray="3 3"
                            vertical={false}
                          />
                          <XAxis
                            dataKey="name"
                            fontSize={10}
                            axisLine={false}
                            tickLine={false}
                          />
                          <YAxis
                            domain={[0, 100]}
                            fontSize={10}
                            axisLine={false}
                            tickLine={false}
                          />
                          <Tooltip />
                          <Bar
                            dataKey="percentage"
                            fill="#16a34a"
                            radius={[4, 4, 0, 0]}
                          />
                        </BarChart>
                      </ResponsiveContainer>
                    ) : (
                      <div className="h-full flex items-center justify-center text-zinc-400 text-xs italic">
                        Memuat data kelas...
                      </div>
                    )}
                  </div>
                  <div className="mt-6">
                    <h4 className="text-[10px] font-black text-red-600 uppercase tracking-widest mb-3">
                      Siswa Kehadiran &lt; 80% (
                      {getMonthYearText(analysisMonth)})
                    </h4>
                    <div className="max-h-48 overflow-y-auto space-y-2">
                      {(() => {
                        const daysCount = effectiveDaysForAnalysisMonth.length;
                        if (daysCount === 0)
                          return (
                            <div className="text-center py-4 text-gray-400 text-[10px] font-bold uppercase italic">
                              Belum ada hari sekolah bulan ini
                            </div>
                          );

                        const filteredStudents = (students || []).filter((s) =>
                          analysisClass ? s.kelas === analysisClass : true,
                        );
                        const lowAttendanceSiswa = filteredStudents.filter(
                          (s) => {
                            const currentMonthString = analysisMonth;
                            const studentAttendance = (attendance || []).filter(
                              (a) =>
                                a.nisn === s.nisn &&
                                a.tanggal.startsWith(currentMonthString),
                            );
                            const hadir = studentAttendance.filter((a) =>
                              ["Hadir", "Izin", "Sakit"].includes(a.status),
                            ).length;
                            const perc = (hadir / daysCount) * 100;
                            return perc < 80;
                          },
                        );

                        if (lowAttendanceSiswa.length === 0)
                          return (
                            <div className="text-center py-4 text-gray-400 text-[10px] font-bold uppercase italic">
                              Semua siswa tertib presensi ✨
                            </div>
                          );

                        return lowAttendanceSiswa.map((s) => {
                          const currentMonthString = analysisMonth;
                          const studentAttendance = (attendance || []).filter(
                            (a) =>
                              a.nisn === s.nisn &&
                              a.tanggal.startsWith(currentMonthString),
                          );
                          const hadir = studentAttendance.filter((a) =>
                            ["Hadir", "Izin", "Sakit"].includes(a.status),
                          ).length;
                          const perc = Math.round((hadir / daysCount) * 100);

                          return (
                            <div
                              key={s.nisn}
                              className="flex justify-between items-center p-2 bg-red-50 rounded-lg border border-red-100"
                            >
                              <div className="flex flex-col">
                                <span className="text-xs font-bold">
                                  {s.nama}
                                </span>
                                <span className="text-[10px] font-medium text-gray-500">
                                  {s.kelas} • {hadir} dari {daysCount} hari
                                  sekolah
                                </span>
                              </div>
                              <span className="text-xs font-black text-red-600">
                                {perc}%
                              </span>
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
                    {session?.role === "Guru" && session?.isWali
                      ? `Analisis Kehadiran Guru Pengajar di ${session.kelas}`
                      : "Analisis Kehadiran Guru"}
                  </h3>
                  <div className="h-64 relative">
                    {teachers.length > 0 ? (
                      <ResponsiveContainer
                        width="100%"
                        height="100%"
                        minWidth={0}
                      >
                        <BarChart
                          data={teachers
                            .filter(
                              (t) => t.nip !== "ADMIN001" && t.role !== "Admin",
                            )
                            .filter((t) => {
                              if (!analysisClass) return true;
                              const scheds = (teachingSchedules || []).filter(
                                (ts) => ts.nip === t.nip,
                              );
                              return scheds.some(
                                (s) => s.kelas === analysisClass,
                              );
                            })
                            .map((t) => {
                              const schedules = (
                                teachingSchedules || []
                              ).filter(
                                (ts) =>
                                  ts.nip === t.nip &&
                                  (analysisClass
                                    ? ts.kelas === analysisClass
                                    : true),
                              );
                              const [y, m] = analysisMonth
                                .split("-")
                                .map(Number);
                              const targetMonth = m - 1;
                              const isCurrentMonth =
                                analysisMonth === currentDate.slice(0, 7);
                              const upToDay = isCurrentMonth
                                ? new Date().getDate()
                                : undefined;
                              const totalTarget = (schedules || []).reduce(
                                (sum, s) => {
                                  return (
                                    sum +
                                    getEffectiveTargetSessionsForSchedule(
                                      s,
                                      y,
                                      targetMonth,
                                      holidays,
                                      settings,
                                      upToDay,
                                    )
                                  );
                                },
                                0,
                              );
                              const currentMonth = analysisMonth;
                              const actual = (teacherAttendance || []).filter(
                                (ta) =>
                                  ta.nip === t.nip &&
                                  (analysisClass
                                    ? ta.kelas === analysisClass
                                    : true) &&
                                  ta.tanggal.startsWith(currentMonth),
                              ).length;
                              return {
                                name: (t.nama || "Guru").split(" ")[0],
                                percentage:
                                  totalTarget > 0
                                    ? Math.round((actual / totalTarget) * 100)
                                    : 0,
                              };
                            })}
                        >
                          <CartesianGrid
                            strokeDasharray="3 3"
                            vertical={false}
                          />
                          <XAxis
                            dataKey="name"
                            fontSize={10}
                            axisLine={false}
                            tickLine={false}
                          />
                          <YAxis
                            domain={[0, 100]}
                            fontSize={10}
                            axisLine={false}
                            tickLine={false}
                          />
                          <Tooltip />
                          <Bar
                            dataKey="percentage"
                            fill="#2563eb"
                            radius={[4, 4, 0, 0]}
                          />
                        </BarChart>
                      </ResponsiveContainer>
                    ) : (
                      <div className="h-full flex items-center justify-center text-zinc-400 text-xs italic">
                        Memuat data guru...
                      </div>
                    )}
                  </div>
                  <div className="mt-6">
                    <h4 className="text-[10px] font-black text-red-600 uppercase tracking-widest mb-3">
                      {session?.role === "Guru" && session?.isWali
                        ? `Performa Guru di ${session.kelas} < 50% (${getMonthYearText(analysisMonth)})`
                        : `Guru Performa < 50% (${getMonthYearText(analysisMonth)})`}
                    </h4>
                    <div className="max-h-48 overflow-y-auto space-y-2">
                      {(teachers || [])
                        .filter(
                          (t) => t.nip !== "ADMIN001" && t.role !== "Admin",
                        )
                        .filter((t) => {
                          if (!analysisClass) return true;
                          const scheds = (teachingSchedules || []).filter(
                            (ts) => ts.nip === t.nip,
                          );
                          return scheds.some((s) => s.kelas === analysisClass);
                        })
                        .map((t) => {
                          const schedules = (teachingSchedules || []).filter(
                            (ts) =>
                              ts.nip === t.nip &&
                              (analysisClass
                                ? ts.kelas === analysisClass
                                : true),
                          );
                          const [y, m] = analysisMonth.split("-").map(Number);
                          const targetMonth = m - 1;
                          const isCurrentMonth =
                            analysisMonth === currentDate.slice(0, 7);
                          const upToDay = isCurrentMonth
                            ? new Date().getDate()
                            : undefined;
                          const totalTarget = (schedules || []).reduce(
                            (sum, s) => {
                              return (
                                sum +
                                getEffectiveTargetSessionsForSchedule(
                                  s,
                                  y,
                                  targetMonth,
                                  holidays,
                                  settings,
                                  upToDay,
                                )
                              );
                            },
                            0,
                          );
                          const currentMonth = analysisMonth;
                          const actual = (teacherAttendance || []).filter(
                            (ta) =>
                              ta.nip === t.nip &&
                              (analysisClass
                                ? ta.kelas === analysisClass
                                : true) &&
                              ta.tanggal.startsWith(currentMonth),
                          ).length;
                          const perc =
                            totalTarget > 0 ? (actual / totalTarget) * 100 : 0;
                          if (perc < 50 && totalTarget > 0) {
                            return (
                              <div
                                key={t.nip}
                                className="flex justify-between items-center p-2 bg-red-50 rounded-lg border border-red-100"
                              >
                                <span className="text-xs font-bold">
                                  {t.nama}
                                </span>
                                <span className="text-xs font-black text-red-600">
                                  {Math.round(perc)}%
                                </span>
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

          {activePanel === "jadwal-mengajar" && (
            <motion.div
              key="jadwal-mengajar"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
            >
              <div className="mb-4 flex flex-wrap gap-3 justify-between items-center text-zinc-900">
                <div className="flex flex-wrap gap-3 items-center">
                  <div className="relative w-64">
                    <Search
                      size={16}
                      className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400"
                    />
                    <input
                      type="text"
                      placeholder="Cari Jadwal..."
                      value={searchTermJadwal}
                      onChange={(e) => {
                        setSearchTermJadwal(e.target.value);
                        setPagination({ ...pagination, jadwal: 0 });
                      }}
                      className="w-full bg-white border border-gray-100 rounded-xl py-2 pl-10 pr-4 text-sm focus:ring-2 focus:ring-green-500"
                    />
                  </div>
                  <select
                    value={filterJadwalClass}
                    onChange={(e) => {
                      setFilterJadwalClass(e.target.value);
                      setPagination({ ...pagination, jadwal: 0 });
                    }}
                    className="bg-white border border-gray-100 rounded-xl px-4 py-2 text-sm font-semibold text-zinc-600 focus:ring-2 focus:ring-green-500"
                  >
                    <option value="">Semua Kelas</option>
                    {classrooms.map((c) => (
                      <option key={c.nama} value={c.nama}>
                        Kelas {c.nama}
                      </option>
                    ))}
                  </select>
                </div>
                <div className="flex gap-2">
                  <select
                    value={pageSize}
                    onChange={(e) => setPageSize(parseInt(e.target.value))}
                    className="bg-zinc-100 text-zinc-600 px-3 py-2 rounded-xl text-[10px] font-black uppercase border-0 focus:ring-0"
                  >
                    {[10, 20, 50, 100].map((v) => (
                      <option key={v} value={v}>
                        Tampil {v}
                      </option>
                    ))}
                  </select>
                  <button
                    onClick={() => {
                      setEditingJadwal({
                        id: "",
                        nip: "",
                        namaGuru: "",
                        kelas: "",
                        hari: "",
                        targetPertemuan: 2,
                      } as any);
                      setTeachingSessions([]);
                      setShowJadwalModal(true);
                    }}
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
                        <th className="px-6 py-4 text-center w-12 text-nowrap">
                          No
                        </th>
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

                        return groupedJadwal
                          .slice(
                            pagination.jadwal,
                            pagination.jadwal + pageSize,
                          )
                          .map((g: any, i) => {
                            const rowKey = `${g.nip}-${g.mapel || "N/A"}`;
                            const isExpanded = expandedJadwal.includes(rowKey);
                            const uniqueClasses = Array.from(
                              new Set(g.sessions.map((s: any) => s.kelas)),
                            )
                              .filter(Boolean)
                              .sort();

                            const totalTarget = g.sessions.reduce(
                              (sum: number, s: any) => {
                                return (
                                  sum +
                                  getEffectiveTargetSessionsForSchedule(
                                    s,
                                    year,
                                    month,
                                    holidays,
                                    settings,
                                  )
                                );
                              },
                              0,
                            );

                            return (
                              <React.Fragment key={rowKey}>
                                <tr
                                  className={isExpanded ? "bg-zinc-50/10" : ""}
                                >
                                  <td className="px-6 py-4 text-center font-bold text-gray-400 text-xs">
                                    {pagination.jadwal + i + 1}
                                  </td>
                                  <td className="px-6 py-4">
                                    <div className="flex items-center gap-2">
                                      <button
                                        onClick={() => {
                                          setExpandedJadwal((prev) =>
                                            prev.includes(rowKey)
                                              ? prev.filter((k) => k !== rowKey)
                                              : [...prev, rowKey],
                                          );
                                        }}
                                        className="p-1 hover:bg-zinc-50 rounded transition-all cursor-pointer"
                                        title={
                                          isExpanded
                                            ? "Sembunyikan kelas"
                                            : "Lihat kelas"
                                        }
                                      >
                                        {isExpanded ? (
                                          <ChevronUp
                                            size={16}
                                            className="text-zinc-500"
                                          />
                                        ) : (
                                          <ChevronDown
                                            size={16}
                                            className="text-zinc-400"
                                          />
                                        )}
                                      </button>
                                      <div>
                                        <p className="font-bold">
                                          {g.namaGuru}
                                        </p>
                                        <p className="text-[9px] font-mono font-bold text-gray-400 uppercase">
                                          {g.nip}
                                        </p>
                                      </div>
                                    </div>
                                  </td>
                                  <td className="px-6 py-4 font-bold text-green-950 text-xs">
                                    {g.mapel || "-"}
                                  </td>
                                  <td className="px-6 py-4">
                                    <div className="flex flex-wrap gap-1 max-w-[180px]">
                                      {uniqueClasses.map((cl: any) => (
                                        <span
                                          key={cl}
                                          className="inline-block bg-green-50 text-green-800 text-[10px] font-black px-2.5 py-0.5 rounded-md border border-green-100/50 uppercase"
                                        >
                                          {cl}
                                        </span>
                                      ))}
                                    </div>
                                  </td>
                                  <td className="px-6 py-4">
                                    <div className="flex flex-col">
                                      <span className="font-black text-green-900">
                                        {totalTarget} Pertemuan
                                      </span>
                                      <span className="text-[9px] font-bold text-gray-400 italic">
                                        Bulan Ini (
                                        {new Date().toLocaleDateString(
                                          "id-ID",
                                          { month: "long" },
                                        )}
                                        )
                                      </span>
                                    </div>
                                  </td>
                                  <td className="px-6 py-4 text-center text-nowrap">
                                    <div className="flex justify-center gap-2">
                                      <button
                                        onClick={() => {
                                          setExpandedJadwal((prev) =>
                                            prev.includes(rowKey)
                                              ? prev.filter((k) => k !== rowKey)
                                              : [...prev, rowKey],
                                          );
                                        }}
                                        className="text-[10px] font-black tracking-wider uppercase text-blue-600 bg-blue-50/50 hover:bg-blue-50 px-3.5 py-2 rounded-xl border border-blue-100/50 transition-all cursor-pointer"
                                      >
                                        {isExpanded ? "Tutup" : "Lihat Sesi"}
                                      </button>
                                      <button
                                        onClick={() => {
                                          setConfirmModal({
                                            show: true,
                                            title: "Hapus Seluruh Jadwal?",
                                            message: `Seluruh (${g.sessions.length}) jadwal mengajar guru ini untuk mapel ${g.mapel} di SEMUA kelas akan dihapus permanen.`,
                                            entityName: `${g.namaGuru} - ${g.mapel}`,
                                            onConfirm: async () => {
                                              for (const s of g.sessions) {
                                                await firestoreService.hapusJadwalMengajar(
                                                  s.id,
                                                );
                                              }
                                              refreshMasterData();
                                            },
                                          });
                                        }}
                                        className="text-red-500 hover:bg-red-50 p-2 border border-transparent hover:border-red-100 rounded-xl transition-all cursor-pointer"
                                        title="Hapus Semua Kelas/Sesi"
                                      >
                                        <Trash2 size={16} />
                                      </button>
                                    </div>
                                  </td>
                                </tr>

                                {isExpanded && (
                                  <tr className="bg-zinc-50/20">
                                    <td
                                      colSpan={6}
                                      className="px-8 py-4 bg-zinc-50/40 border-t border-b border-gray-100"
                                    >
                                      <div className="text-[9px] font-black text-zinc-400 uppercase tracking-widest mb-3">
                                        Detail Mengajar per Kelas:
                                      </div>
                                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                                        {uniqueClasses.map((cls: any) => {
                                          const classSessions =
                                            g.sessions.filter(
                                              (s: any) => s.kelas === cls,
                                            );
                                          const totalClassTarget =
                                            classSessions.reduce(
                                              (sum: number, s: any) => {
                                                return (
                                                  sum +
                                                  getEffectiveTargetSessionsForSchedule(
                                                    s,
                                                    year,
                                                    month,
                                                    holidays,
                                                    settings,
                                                  )
                                                );
                                              },
                                              0,
                                            );

                                          return (
                                            <div
                                              key={cls}
                                              className="bg-white p-4 rounded-2xl border border-gray-100 flex justify-between items-center shadow-sm hover:border-zinc-200 transition-all"
                                            >
                                              <div>
                                                <div className="flex items-center gap-2 mb-2">
                                                  <span className="text-xs font-black text-green-950 bg-green-50 border border-green-100/50 px-2.5 py-1 rounded-lg">
                                                    Kelas {cls}
                                                  </span>
                                                  <span className="text-[10px] text-zinc-400 font-bold">
                                                    {totalClassTarget} Sesi /
                                                    Bulan
                                                  </span>
                                                </div>
                                                <div className="space-y-1">
                                                  {classSessions.map(
                                                    (cs: any, idx: number) => (
                                                      <div
                                                        key={idx}
                                                        className="flex flex-wrap items-center gap-1.5 text-xs text-zinc-600 font-medium leading-tight"
                                                      >
                                                        <span className="w-1.5 h-1.5 rounded-full bg-zinc-300"></span>
                                                        <span className="font-bold text-zinc-850">
                                                          {cs.hari}
                                                        </span>
                                                        {cs.jps &&
                                                          cs.jps.length > 0 && (
                                                            <span className="text-zinc-400 text-[10px] font-mono font-bold bg-zinc-50 px-1.5 py-0.5 rounded border border-zinc-100">
                                                              (
                                                              {cs.jps
                                                                .map(
                                                                  (j: number) =>
                                                                    `JP ${j}`,
                                                                )
                                                                .join(", ")}
                                                              )
                                                            </span>
                                                          )}
                                                        <span className="text-[10px] text-zinc-400 font-bold italic">
                                                          ({cs.targetPertemuan}{" "}
                                                          target/mgu)
                                                        </span>
                                                      </div>
                                                    ),
                                                  )}
                                                </div>
                                              </div>

                                              <div className="flex gap-1.5">
                                                <button
                                                  onClick={() => {
                                                    setEditingJadwal({
                                                      id: "",
                                                      nip: g.nip,
                                                      namaGuru: g.namaGuru,
                                                      mapel: g.mapel,
                                                      kelas: cls,
                                                    } as any);
                                                    setTeachingSessions(
                                                      classSessions.map(
                                                        (s: any) => ({
                                                          kelas: s.kelas,
                                                          hari: s.hari,
                                                          target:
                                                            s.targetPertemuan,
                                                          jps: s.jps || [],
                                                        }),
                                                      ),
                                                    );
                                                    setShowJadwalModal(true);
                                                  }}
                                                  className="text-blue-500 hover:bg-blue-50 p-2 border border-transparent hover:border-blue-100 rounded-xl transition-all h-8 w-8 flex items-center justify-center cursor-pointer"
                                                  title="Edit Jadwal Kelas Ini"
                                                >
                                                  <Edit size={14} />
                                                </button>
                                                <button
                                                  onClick={() => {
                                                    setConfirmModal({
                                                      show: true,
                                                      title:
                                                        "Hapus Jam Mengajar Kelas?",
                                                      message: `Seluruh jam mengajar guru untuk mapel ini di kelas ${cls} akan dihapus.`,
                                                      entityName: `${g.namaGuru} - Kelas ${cls} (${g.mapel})`,
                                                      onConfirm: async () => {
                                                        for (const s of classSessions) {
                                                          await firestoreService.hapusJadwalMengajar(
                                                            s.id,
                                                          );
                                                        }
                                                      },
                                                    });
                                                  }}
                                                  className="text-red-500 hover:bg-red-50 p-2 border border-transparent hover:border-red-100 rounded-xl transition-all h-8 w-8 flex items-center justify-center cursor-pointer"
                                                  title="Hapus Jadwal Kelas Ini"
                                                >
                                                  <Trash2 size={14} />
                                                </button>
                                              </div>
                                            </div>
                                          );
                                        })}
                                      </div>
                                    </td>
                                  </tr>
                                )}
                              </React.Fragment>
                            );
                          });
                      })()}
                    </tbody>
                  </table>
                </div>
              </div>
              {groupedJadwal.length > pageSize && (
                <div className="mt-4 flex justify-end items-center gap-4 bg-white p-4 rounded-2xl border border-gray-100 shadow-sm text-zinc-900">
                  <span className="text-[10px] font-black text-gray-400 uppercase tracking-widest">
                    Halaman {Math.floor(pagination.jadwal / pageSize) + 1}
                  </span>
                  <div className="flex gap-2">
                    <button
                      disabled={pagination.jadwal === 0}
                      onClick={() =>
                        setPagination({
                          ...pagination,
                          jadwal: pagination.jadwal - pageSize,
                        })
                      }
                      className="p-2 rounded-lg bg-white border border-gray-100 text-gray-400 disabled:opacity-30 hover:bg-gray-50 transition-all"
                    >
                      <ChevronLeft size={16} />
                    </button>
                    <button
                      disabled={
                        pagination.jadwal + pageSize >= groupedJadwal.length
                      }
                      onClick={() =>
                        setPagination({
                          ...pagination,
                          jadwal: pagination.jadwal + pageSize,
                        })
                      }
                      className="p-2 rounded-lg bg-white border border-gray-100 text-gray-400 disabled:opacity-30 hover:bg-gray-50 transition-all"
                    >
                      <ChevronRight size={16} />
                    </button>
                  </div>
                </div>
              )}
            </motion.div>
          )}

          {activePanel === "guru" && (
            <motion.div
              key="guru"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
            >
              <div className="mb-4 flex flex-wrap gap-3 justify-between items-center text-zinc-900">
                <div className="relative w-72">
                  <Search
                    size={16}
                    className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400"
                  />
                  <input
                    type="text"
                    placeholder="Cari Guru..."
                    value={searchTermGuru}
                    onChange={(e) => {
                      setSearchTermGuru(e.target.value);
                      setPagination({ ...pagination, guru: 0 });
                    }}
                    className="w-full bg-white border border-gray-100 rounded-xl py-2 pl-10 pr-4 text-sm focus:ring-2 focus:ring-green-500"
                  />
                </div>
                <div className="flex gap-2 text-zinc-900">
                  <select
                    value={pageSize}
                    onChange={(e) => setPageSize(parseInt(e.target.value))}
                    className="bg-zinc-100 text-zinc-600 px-3 py-2 rounded-xl text-[10px] font-black uppercase border-0 focus:ring-0"
                  >
                    {[10, 20, 50, 100].map((v) => (
                      <option key={v} value={v}>
                        Tampil {v}
                      </option>
                    ))}
                  </select>
                  <button
                    onClick={() => downloadTemplate("Guru")}
                    className="bg-zinc-100 text-zinc-600 px-4 py-2 rounded-xl text-[10px] font-black tracking-widest uppercase flex items-center gap-2 hover:bg-zinc-200 transition-all"
                  >
                    <Download size={14} /> Template
                  </button>
                  <button
                    onClick={() => handleImportExcel("Guru")}
                    className="bg-blue-100 text-blue-600 px-4 py-2 rounded-xl text-[10px] font-black tracking-widest uppercase flex items-center gap-2 hover:bg-blue-200 transition-all"
                  >
                    <Upload size={14} /> Import
                  </button>
                  <button
                    onClick={() => {
                      setEditingGuru({
                        nip: "",
                        nama: "",
                        jabatan: "Guru",
                        kelas: "",
                        user: "",
                        pass: "",
                        role: "Guru",
                      });
                      setShowGuruModal(true);
                    }}
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
                        <th className="px-6 py-4 text-center w-12 text-nowrap">
                          No
                        </th>
                        <th className="px-6 py-4">Foto / NIP</th>
                        <th className="px-6 py-4">Nama Guru</th>
                        <th className="px-6 py-4">Jabatan</th>
                        <th className="px-6 py-4">Wali Kelas</th>
                        <th className="px-6 py-4">Username</th>
                        <th className="px-6 py-4">Aksi</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-50">
                      {filteredGuru
                        .slice(pagination.guru, pagination.guru + pageSize)
                        .map((g, i) => (
                          <tr
                            key={g.nip}
                            className="hover:bg-gray-50 transition-colors"
                          >
                            <td className="px-6 py-4 text-center font-bold text-gray-400 text-xs">
                              {pagination.guru + i + 1}
                            </td>
                            <td className="px-6 py-4">
                              <div className="w-10 h-10 rounded-lg bg-gray-100 flex items-center justify-center overflow-hidden border">
                                {g.foto ? (
                                  <img
                                    src={g.foto}
                                    className="w-full h-full object-cover"
                                  />
                                ) : (
                                  <User size={16} className="text-gray-300" />
                                )}
                              </div>
                            </td>
                            <td className="px-6 py-4">
                              <p className="font-bold">{g.nama}</p>
                              <p className="text-[10px] font-mono font-bold text-gray-400">
                                NIP: {g.nip}
                              </p>
                            </td>
                            <td className="px-6 py-4">
                              <span className="bg-zinc-100 text-zinc-600 px-2 py-0.5 rounded text-[9px] font-black uppercase">
                                {g.jabatan || "Guru"}
                              </span>
                            </td>
                            <td className="px-6 py-4">
                              <span className="text-green-800 text-[10px] font-bold">
                                {g.kelas || "-"}
                              </span>
                            </td>
                            <td className="px-6 py-4 text-gray-600 font-medium">
                              {g.user}
                            </td>
                            <td className="px-6 py-4">
                              <div className="flex gap-2">
                                <button
                                  onClick={() => {
                                    setEditingGuru({
                                      nip: g.nip,
                                      nama: g.nama,
                                      jabatan: g.jabatan || "Guru",
                                      kelas: g.kelas || "",
                                      user: g.user || g.nip || "",
                                      pass: g.pass || "123456",
                                      role: g.role || "Guru",
                                      foto: g.foto || "",
                                    });
                                    setShowGuruModal(true);
                                  }}
                                  className="text-blue-500 hover:text-blue-800"
                                >
                                  <Edit size={16} />
                                </button>
                                <button
                                  onClick={() =>
                                    handleDeleteGuru(g.nip, g.nama)
                                  }
                                  className="text-red-500 hover:text-red-800"
                                >
                                  <Trash2 size={16} />
                                </button>
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
                  <span className="text-[10px] font-black text-gray-400 uppercase tracking-widest">
                    Halaman {Math.floor(pagination.guru / pageSize) + 1}
                  </span>
                  <div className="flex gap-2">
                    <button
                      disabled={pagination.guru === 0}
                      onClick={() =>
                        setPagination({
                          ...pagination,
                          guru: pagination.guru - pageSize,
                        })
                      }
                      className="p-2 rounded-lg bg-white border border-gray-100 text-gray-400 disabled:opacity-30 hover:bg-gray-50 transition-all"
                    >
                      <ChevronLeft size={16} />
                    </button>
                    <button
                      disabled={
                        pagination.guru + pageSize >= filteredGuru.length
                      }
                      onClick={() =>
                        setPagination({
                          ...pagination,
                          guru: pagination.guru + pageSize,
                        })
                      }
                      className="p-2 rounded-lg bg-white border border-gray-100 text-gray-400 disabled:opacity-30 hover:bg-gray-50 transition-all"
                    >
                      <ChevronRight size={16} />
                    </button>
                  </div>
                </div>
              )}
            </motion.div>
          )}

          {activePanel === "kelas" && (
            <motion.div
              key="kelas"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
            >
              <div className="mb-4 flex flex-wrap gap-3 justify-between items-center text-zinc-900">
                <div className="flex flex-wrap gap-3 items-center">
                  <div className="relative w-64">
                    <Search
                      size={16}
                      className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400"
                    />
                    <input
                      type="text"
                      placeholder="Cari Kelas..."
                      value={searchTermKelas}
                      onChange={(e) => {
                        setSearchTermKelas(e.target.value);
                        setPagination({ ...pagination, kelas: 0 });
                      }}
                      className="w-full bg-white border border-gray-100 rounded-xl py-2 pl-10 pr-4 text-sm focus:ring-2 focus:ring-green-500"
                    />
                  </div>
                  <select
                    value={pageSize}
                    onChange={(e) => setPageSize(parseInt(e.target.value))}
                    className="bg-zinc-100 text-zinc-600 px-3 py-2 rounded-xl text-[10px] font-black uppercase border-0 focus:ring-0"
                  >
                    {[10, 20, 50, 100].map((v) => (
                      <option key={v} value={v}>
                        Tampil {v}
                      </option>
                    ))}
                  </select>
                </div>
                <button
                  onClick={() => {
                    setEditingKelas({ nama: "", wali: "", jumlah: 0 });
                    setShowKelasModal(true);
                  }}
                  className="bg-green-800 text-white px-4 py-2 rounded-xl text-sm font-bold flex items-center gap-2"
                >
                  <Plus size={16} /> Tambah Kelas
                </button>
              </div>
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4 text-zinc-900">
                {filteredKelas
                  .slice(pagination.kelas, pagination.kelas + pageSize)
                  .map((k, i) => (
                    <div
                      key={i}
                      className="bg-white p-6 rounded-2xl shadow-sm border border-gray-100 flex flex-col justify-between"
                    >
                      <div>
                        <div className="flex justify-between items-start mb-4">
                          <div className="flex flex-col">
                            <span className="text-[10px] font-black text-gray-400">
                              #{pagination.kelas + i + 1}
                            </span>
                            <h3 className="text-2xl font-black text-green-900">
                              {k.nama}
                            </h3>
                          </div>
                          <div className="bg-green-100 p-2 rounded-xl text-green-700">
                            <TableIcon size={20} />
                          </div>
                        </div>
                        <p className="text-sm font-bold text-gray-400 uppercase tracking-tight">
                          Wali Kelas
                        </p>
                        <p className="font-bold text-gray-800 mb-4">{k.wali}</p>

                        <div className="flex items-center justify-between p-3">
                          <span className="text-xs font-bold text-gray-400">
                            KUOTA
                          </span>
                          <span className="text-sm font-black">
                            {k.jumlah} Siswa
                          </span>
                        </div>

                        <div className="mt-4 flex flex-col items-center p-3 rounded-2xl">
                          <p className="text-[9px] font-black text-gray-400 uppercase mb-2">
                            Barcode Absensi Guru
                          </p>
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
                          style={{
                            left: "-5000px",
                            top: "-5000px",
                            zIndex: -1000,
                          }}
                        >
                          <div
                            style={{
                              padding: "50px",
                              backgroundColor: "#ffffff",
                              textAlign: "center",
                              width: "600px",
                              display: "flex",
                              flexDirection: "column",
                              alignItems: "center",
                              justifyContent: "center",
                            }}
                          >
                            <h2
                              style={{
                                color: "#14532d",
                                fontSize: "36px",
                                marginBottom: "30px",
                                fontWeight: "900",
                                lineHeight: "1.2",
                                textTransform: "uppercase",
                              }}
                            >
                              ABSENSI MENGAJAR
                              <br />
                              KELAS {k.nama}
                            </h2>
                            <div
                              style={{
                                backgroundColor: "#ffffff",
                                padding: "20px",
                                borderRadius: "30px",
                                boxShadow: "0 10px 25px rgba(0,0,0,0.1)",
                                border: "1px solid #f0f0f0",
                                display: "flex",
                                alignItems: "center",
                                justifyContent: "center",
                              }}
                            >
                              <QRCodeSVG
                                value={k.nama}
                                size={400}
                                level="H"
                                includeMargin={false}
                              />
                            </div>
                            <p
                              style={{
                                color: "#14532d",
                                fontSize: "28px",
                                marginTop: "40px",
                                fontWeight: "900",
                                letterSpacing: "0.1em",
                              }}
                            >
                              SIGAP MTsN 2 BOMBANA
                            </p>
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
                        <button
                          onClick={() => {
                            setEditingKelas({ ...k, oldNama: k.nama });
                            setShowKelasModal(true);
                          }}
                          className="bg-gray-100 text-gray-600 p-2 rounded-lg hover:bg-gray-200"
                        >
                          <Edit size={14} />
                        </button>
                        <button
                          onClick={() => handleDeleteKelas(k.nama)}
                          className="bg-red-50 text-red-600 p-2 rounded-lg hover:bg-red-100"
                        >
                          <Trash2 size={14} />
                        </button>
                      </div>
                    </div>
                  ))}
              </div>
              {filteredKelas.length > pageSize && (
                <div className="mt-4 flex justify-end items-center gap-4 bg-white p-4 rounded-2xl border border-gray-100 shadow-sm text-zinc-900">
                  <span className="text-[10px] font-black text-gray-400 uppercase tracking-widest">
                    Halaman {Math.floor(pagination.kelas / pageSize) + 1}
                  </span>
                  <div className="flex gap-2">
                    <button
                      disabled={pagination.kelas === 0}
                      onClick={() =>
                        setPagination({
                          ...pagination,
                          kelas: pagination.kelas - pageSize,
                        })
                      }
                      className="p-2 rounded-lg bg-white border border-gray-100 text-gray-400 disabled:opacity-30 hover:bg-gray-50 transition-all"
                    >
                      <ChevronLeft size={16} />
                    </button>
                    <button
                      disabled={
                        pagination.kelas + pageSize >= filteredKelas.length
                      }
                      onClick={() =>
                        setPagination({
                          ...pagination,
                          kelas: pagination.kelas + pageSize,
                        })
                      }
                      className="p-2 rounded-lg bg-white border border-gray-100 text-gray-400 disabled:opacity-30 hover:bg-gray-50 transition-all"
                    >
                      <ChevronRight size={16} />
                    </button>
                  </div>
                </div>
              )}
            </motion.div>
          )}

          {activePanel === "absensi-wali" && (
            <motion.div
              key="absensi-wali"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
            >
              <div className="grid grid-cols-1 xl:grid-cols-3 gap-6">
                <div className="xl:col-span-2 space-y-6">
                  {(() => {
                    const classStudents = students.filter(
                      (s) => s.kelas === session?.kelas,
                    );
                    const classAttendanceToday = attendance.filter(
                      (a) =>
                        a.tanggal === waliFilterTanggal &&
                        classStudents.some((s) => s.nisn === a.nisn),
                    );

                    const countHadir = classAttendanceToday.filter(
                      (a) => a.status === "Hadir",
                    ).length;
                    const countTerlambat = classAttendanceToday.filter(
                      (a) => a.status === "Hadir" && a.terlambat > 0,
                    ).length;
                    const countSakit = classAttendanceToday.filter(
                      (a) => a.status === "Sakit",
                    ).length;
                    const countIzin = classAttendanceToday.filter(
                      (a) => a.status === "Izin",
                    ).length;

                    const isHoliday =
                      holidays.some((h) => h.tanggal === waliFilterTanggal) ||
                      new Date(waliFilterTanggal).getDay() === 0;
                    const recordedAlfa = classAttendanceToday.filter(
                      (a) => a.status === "Alfa",
                    ).length;
                    const attendedNisns = new Set(
                      classAttendanceToday.map((a) => a.nisn),
                    );
                    const notYetCheckedIn =
                      classStudents.length - attendedNisns.size;
                    const countAlfa = isHoliday
                      ? 0
                      : recordedAlfa + notYetCheckedIn;

                    return (
                      <div className="grid grid-cols-2 md:grid-cols-5 gap-3">
                        <div className="bg-white p-4 rounded-xl border border-gray-100 shadow-sm flex items-center gap-3">
                          <div className="p-2.5 rounded-lg bg-green-50 text-green-600 shrink-0">
                            <CheckCircle2 size={18} />
                          </div>
                          <div className="min-w-0">
                            <p className="text-[9px] text-gray-400 font-bold uppercase tracking-wider truncate">
                              Hadir
                            </p>
                            <p className="text-sm font-black text-gray-950 mt-0.5">
                              {countHadir}{" "}
                              <span className="text-gray-400 text-[10px] font-normal">
                                /{classStudents.length}
                              </span>
                            </p>
                          </div>
                        </div>

                        <div className="bg-white p-4 rounded-xl border border-gray-100 shadow-sm flex items-center gap-3">
                          <div className="p-2.5 rounded-lg bg-amber-50 text-amber-600 shrink-0">
                            <Clock size={18} />
                          </div>
                          <div className="min-w-0">
                            <p className="text-[9px] text-gray-400 font-bold uppercase tracking-wider truncate">
                              Terlambat
                            </p>
                            <p className="text-sm font-black text-gray-950 mt-0.5">
                              {countTerlambat}
                            </p>
                          </div>
                        </div>

                        <div className="bg-white p-4 rounded-xl border border-gray-100 shadow-sm flex items-center gap-3">
                          <div className="p-2.5 rounded-lg bg-red-50 text-red-600 shrink-0">
                            <Heart size={18} />
                          </div>
                          <div className="min-w-0">
                            <p className="text-[9px] text-gray-400 font-bold uppercase tracking-wider truncate">
                              Sakit
                            </p>
                            <p className="text-sm font-black text-gray-950 mt-0.5">
                              {countSakit}
                            </p>
                          </div>
                        </div>

                        <div className="bg-white p-4 rounded-xl border border-gray-100 shadow-sm flex items-center gap-3">
                          <div className="p-2.5 rounded-lg bg-indigo-50 text-indigo-600 shrink-0">
                            <Info size={18} />
                          </div>
                          <div className="min-w-0">
                            <p className="text-[9px] text-gray-400 font-bold uppercase tracking-wider truncate">
                              Izin
                            </p>
                            <p className="text-sm font-black text-gray-950 mt-0.5">
                              {countIzin}
                            </p>
                          </div>
                        </div>

                        <div className="bg-white p-4 rounded-xl border border-gray-100 shadow-sm flex items-center gap-3 col-span-2 md:col-span-1">
                          <div className="p-2.5 rounded-lg bg-rose-50 text-rose-600 shrink-0">
                            <XCircle size={18} />
                          </div>
                          <div className="min-w-0">
                            <p className="text-[9px] text-gray-400 font-bold uppercase tracking-wider truncate">
                              Alfa
                            </p>
                            <p className="text-sm font-black text-gray-950 mt-0.5">
                              {countAlfa}
                            </p>
                          </div>
                        </div>
                      </div>
                    );
                  })()}

                  <div className="bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden">
                    <div className="p-6 bg-green-50 border-b border-green-100 flex flex-col sm:flex-row gap-4 justify-between sm:items-center">
                      <div>
                        <h2 className="text-sm font-black text-green-900 uppercase tracking-widest">
                          Riwayat Kehadiran Siswa ({session?.kelas})
                        </h2>
                        <p className="text-[10px] text-green-700 font-bold mt-0.5">
                          Hasil absensi tanggal{" "}
                          {formatIndoDate(waliFilterTanggal)}.
                        </p>
                      </div>
                      <div className="flex items-center gap-2">
                        <select
                          value={pageSize}
                          onChange={(e) => {
                            setPageSize(parseInt(e.target.value));
                            setWaliPagination(0);
                          }}
                          className="bg-zinc-100 text-zinc-600 px-3 py-2 rounded-xl text-[10px] font-black uppercase border-0 focus:ring-0"
                        >
                          {[10, 20, 50, 100].map((v) => (
                            <option key={v} value={v}>
                              Tampil {v}
                            </option>
                          ))}
                        </select>
                        <div className="bg-white px-3 py-1.5 rounded-xl border border-green-200 text-[10px] font-black">
                          {formatIndoDate(waliFilterTanggal)}
                        </div>
                      </div>
                    </div>

                    {/* Filter Bar */}
                    <div className="p-4 bg-zinc-50 border-b border-zinc-100 grid grid-cols-1 md:grid-cols-3 gap-3">
                      <div>
                        <label className="text-[9px] font-black text-gray-400 uppercase tracking-widest ml-1">
                          Pencarian Nama
                        </label>
                        <div className="relative mt-1">
                          <Search
                            size={12}
                            className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-300"
                          />
                          <input
                            type="text"
                            placeholder="Cari nama siswa..."
                            value={waliFilterNama}
                            onChange={(e) => {
                              setWaliFilterNama(e.target.value);
                              setWaliPagination(0);
                            }}
                            className="w-full bg-white border border-zinc-200 rounded-xl px-9 py-1.5 text-xs font-bold focus:ring-1 focus:ring-green-500 focus:outline-none"
                          />
                        </div>
                      </div>

                      <div>
                        <label className="text-[9px] font-black text-gray-400 uppercase tracking-widest ml-1">
                          Status Kehadiran
                        </label>
                        <select
                          value={waliFilterStatus}
                          onChange={(e) => {
                            setWaliFilterStatus(e.target.value);
                            setWaliPagination(0);
                          }}
                          className="w-full bg-white border border-zinc-200 rounded-xl px-3 py-1.5 text-xs font-bold mt-1 focus:ring-1 focus:ring-green-500 focus:outline-none focus:border-green-500"
                        >
                          <option value="">Semua Status</option>
                          <option value="Hadir">Hadir</option>
                          <option value="Sakit">Sakit</option>
                          <option value="Izin">Izin</option>
                          <option value="Alfa">Alfa</option>
                          <option value="Belum Absen">Belum Absen</option>
                        </select>
                      </div>

                      <div>
                        <label className="text-[9px] font-black text-gray-400 uppercase tracking-widest ml-1">
                          Pilih Tanggal
                        </label>
                        <input
                          type="date"
                          value={waliFilterTanggal}
                          onChange={(e) => {
                            setWaliFilterTanggal(e.target.value);
                            setWaliPagination(0);
                          }}
                          className="w-full bg-white border border-zinc-200 rounded-xl px-3 py-1.5 text-xs font-bold mt-1 focus:ring-1 focus:ring-green-500 focus:outline-none"
                        />
                      </div>
                    </div>

                    <div className="overflow-x-auto">
                      <table className="w-full text-left text-sm min-w-[600px]">
                        <thead className="bg-gray-50 text-gray-400 uppercase text-[10px] font-black tracking-widest border-b border-gray-100">
                          <tr>
                            <th className="px-6 py-4 text-center w-12 text-nowrap">
                              No
                            </th>
                            <th className="px-6 py-4">Siswa</th>
                            <th className="px-6 py-4">Status</th>
                            <th className="px-6 py-4">Keterangan</th>
                            <th className="px-6 py-4 text-center">Update</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-gray-50">
                          {(() => {
                            const filteredWali = students
                              .filter((s) => s.kelas === session?.kelas)
                              .filter((s) => {
                                if (
                                  waliFilterNama &&
                                  !s.nama
                                    .toLowerCase()
                                    .includes(waliFilterNama.toLowerCase())
                                ) {
                                  return false;
                                }
                                const current = attendance.find(
                                  (a) =>
                                    a.nisn === s.nisn &&
                                    a.tanggal === waliFilterTanggal,
                                );
                                const statusValue =
                                  current?.status || "Belum Absen";

                                if (waliFilterStatus) {
                                  if (waliFilterStatus === "Belum Absen") {
                                    return !current;
                                  }
                                  return current?.status === waliFilterStatus;
                                }
                                return true;
                              })
                              .sort((a, b) => a.nama.localeCompare(b.nama));

                            if (filteredWali.length === 0) {
                              return (
                                <tr>
                                  <td
                                    colSpan={5}
                                    className="px-6 py-12 text-center text-gray-400 font-bold uppercase tracking-widest text-[10px]"
                                  >
                                    Tidak ada data siswa yang cocok dengan
                                    filter.
                                  </td>
                                </tr>
                              );
                            }

                            return filteredWali
                              .slice(waliPagination, waliPagination + pageSize)
                              .map((s, i) => {
                                const current = attendance.find(
                                  (a) =>
                                    a.nisn === s.nisn &&
                                    a.tanggal === waliFilterTanggal,
                                );
                                return (
                                  <tr
                                    key={s.nisn}
                                    className="hover:bg-gray-50 transition-colors"
                                  >
                                    <td className="px-6 py-4 text-center font-bold text-gray-400 text-xs">
                                      {waliPagination + i + 1}
                                    </td>
                                    <td className="px-6 py-4">
                                      <p className="font-bold text-zinc-800">
                                        {s.nama}
                                      </p>
                                      <p className="text-[10px] text-gray-400 font-mono">
                                        {s.nisn}
                                      </p>
                                    </td>
                                    <td className="px-6 py-4">
                                      <span
                                        className={`px-3 py-1 rounded-full text-[10px] font-black ${
                                          current?.status === "Hadir"
                                            ? "bg-green-100 text-green-700"
                                            : current?.status === "Sakit"
                                              ? "bg-amber-100 text-amber-700"
                                              : current?.status === "Izin"
                                                ? "bg-indigo-100 text-indigo-700"
                                                : current?.status === "Alfa"
                                                  ? "bg-rose-100 text-rose-700"
                                                  : "bg-gray-100 text-gray-400"
                                        }`}
                                      >
                                        {current?.status || "BELUM ABSEN"}
                                      </span>
                                    </td>
                                    <td className="px-6 py-4 text-[11px] text-gray-500 italic max-w-[200px] truncate">
                                      {current?.keterangan || "-"}
                                    </td>
                                    <td className="px-6 py-4 text-center">
                                      <button
                                        onClick={() =>
                                          setSiswaAbsensiManual({
                                            nisn: s.nisn,
                                            status: current?.status || "Hadir",
                                            ket: current?.keterangan || "",
                                            tanggal: waliFilterTanggal,
                                          })
                                        }
                                        className="p-2 border border-zinc-100 rounded-lg hover:bg-zinc-50 text-green-600"
                                      >
                                        <Edit size={14} />
                                      </button>
                                    </td>
                                  </tr>
                                );
                              });
                          })()}
                        </tbody>
                      </table>
                    </div>

                    {/* Pagination for Wali Kelas */}
                    {(() => {
                      const totalFilteredWali = students
                        .filter((s) => s.kelas === session?.kelas)
                        .filter((s) => {
                          if (
                            waliFilterNama &&
                            !s.nama
                              .toLowerCase()
                              .includes(waliFilterNama.toLowerCase())
                          ) {
                            return false;
                          }
                          const current = attendance.find(
                            (a) =>
                              a.nisn === s.nisn &&
                              a.tanggal === waliFilterTanggal,
                          );
                          if (waliFilterStatus) {
                            if (waliFilterStatus === "Belum Absen") {
                              return !current;
                            }
                            return current?.status === waliFilterStatus;
                          }
                          return true;
                        }).length;

                      if (totalFilteredWali > pageSize) {
                        return (
                          <div className="mt-2 flex justify-end items-center gap-4 bg-zinc-50 p-4 border-t border-gray-100 text-zinc-900 px-6">
                            <span className="text-[10px] font-black text-gray-400 uppercase tracking-widest">
                              Halaman{" "}
                              {Math.floor(waliPagination / pageSize) + 1} S/D{" "}
                              {Math.ceil(totalFilteredWali / pageSize)} (
                              {totalFilteredWali} Siswa)
                            </span>
                            <div className="flex gap-2">
                              <button
                                disabled={waliPagination === 0}
                                onClick={() =>
                                  setWaliPagination(
                                    Math.max(0, waliPagination - pageSize),
                                  )
                                }
                                className="p-2 rounded-lg bg-white border border-gray-100 text-gray-400 disabled:opacity-30 hover:bg-gray-50 transition-all"
                              >
                                <ChevronLeft size={16} />
                              </button>
                              <button
                                disabled={
                                  waliPagination + pageSize >= totalFilteredWali
                                }
                                onClick={() =>
                                  setWaliPagination(waliPagination + pageSize)
                                }
                                className="p-2 rounded-lg bg-white border border-gray-100 text-gray-400 disabled:opacity-30 hover:bg-gray-50 transition-all"
                              >
                                <ChevronRight size={16} />
                              </button>
                            </div>
                          </div>
                        );
                      }
                      return null;
                    })()}
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
                        <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest ml-1">
                          Pilih Tanggal
                        </label>
                        <input
                          type="date"
                          value={siswaAbsensiManual.tanggal}
                          onChange={(e) =>
                            setSiswaAbsensiManual({
                              ...siswaAbsensiManual,
                              tanggal: e.target.value,
                            })
                          }
                          className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold text-sm"
                        />
                      </div>
                      <div>
                        <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest ml-1">
                          Pilih Siswa
                        </label>
                        <select
                          value={siswaAbsensiManual.nisn}
                          onChange={(e) => {
                            const nisn = e.target.value;
                            const current = attendance.find(
                              (a) =>
                                a.nisn === nisn &&
                                a.tanggal === siswaAbsensiManual.tanggal,
                            );
                            setSiswaAbsensiManual({
                              ...siswaAbsensiManual,
                              nisn,
                              status: current?.status || "Hadir",
                              ket: current?.keterangan || "",
                            });
                          }}
                          className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold text-sm"
                        >
                          <option value="">-- Pilih Siswa --</option>
                          {students
                            .filter((s) => s.kelas === session?.kelas)
                            .map((s) => (
                              <option key={s.nisn} value={s.nisn}>
                                {s.nama}
                              </option>
                            ))}
                        </select>
                      </div>
                      <div className="grid grid-cols-2 gap-2">
                        {["Hadir", "Sakit", "Izin", "Alfa"].map((stat) => (
                          <button
                            key={stat}
                            onClick={() =>
                              setSiswaAbsensiManual({
                                ...siswaAbsensiManual,
                                status: stat,
                              })
                            }
                            className={`py-3 rounded-xl text-[10px] font-black uppercase transition-all ${siswaAbsensiManual.status === stat ? "bg-green-800 text-white shadow-lg scale-105" : "bg-zinc-100 text-zinc-400 hover:bg-zinc-200"}`}
                          >
                            {stat}
                          </button>
                        ))}
                      </div>
                      <div>
                        <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest ml-1">
                          Keterangan / Alasan
                        </label>
                        <textarea
                          rows={3}
                          value={siswaAbsensiManual.ket}
                          onChange={(e) =>
                            setSiswaAbsensiManual({
                              ...siswaAbsensiManual,
                              ket: e.target.value,
                            })
                          }
                          className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-3 font-bold text-sm"
                          placeholder="Penyebab tidak hadir..."
                        />
                      </div>
                      <button
                        disabled={!siswaAbsensiManual.nisn}
                        onClick={async () => {
                          const student = students.find(
                            (s) => s.nisn === siswaAbsensiManual.nisn,
                          );
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
                              jam: new Date().toLocaleTimeString("id-ID", {
                                hour: "2-digit",
                                minute: "2-digit",
                                hour12: false,
                              }),
                            });
                            alert("Absensi berhasil disimpan!");
                          } catch (e) {
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

          {activePanel === "absensi-umum" && (
            <motion.div
              key="absensi-umum"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
            >
              <div className="mb-4 flex flex-wrap gap-3 justify-between items-center">
                <div>
                  <h2 className="text-xl font-bold flex items-center gap-2 text-zinc-900">
                    <ClipboardList className="text-green-700" /> Log Presensi
                    Siswa
                  </h2>
                  <p className="text-xs text-zinc-400 font-medium">
                    Manajemen riwayat dan data presensi harian siswa
                  </p>
                </div>
                {session?.role === "Admin" && (
                  <button
                    onClick={() => {
                      setResetAbsensiTarget("Siswa");
                    }}
                    className="bg-red-600 hover:bg-red-700 text-white px-4 py-2.5 rounded-xl text-[10px] font-black tracking-widest uppercase flex items-center gap-2 transition-all shadow-sm"
                  >
                    <Trash2 size={14} /> Reset Absensi Siswa
                  </button>
                )}
              </div>

              <div className="mb-6 flex flex-col md:flex-row gap-4 bg-white p-6 rounded-3xl border border-gray-100 shadow-sm">
                <div className="flex-1">
                  <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest ml-1">
                    Filter Kelas
                  </label>
                  <select
                    value={filterClassAbsensi}
                    onChange={(e) => setFilterClassAbsensi(e.target.value)}
                    className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-2 font-bold text-sm mt-1"
                  >
                    <option value="">Semua Kelas</option>
                    {classrooms.map((c) => (
                      <option key={c.nama} value={c.nama}>
                        {c.nama}
                      </option>
                    ))}
                  </select>
                </div>
                <div className="flex-1">
                  <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest ml-1">
                    Cari Nama
                  </label>
                  <div className="relative mt-1">
                    <Search
                      size={14}
                      className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-300"
                    />
                    <input
                      type="text"
                      placeholder="Ketik nama siswa..."
                      value={filterNamaAbsensi}
                      onChange={(e) => setFilterNamaAbsensi(e.target.value)}
                      className="w-full bg-zinc-50 border-0 rounded-xl px-10 py-2 font-bold text-sm"
                    />
                  </div>
                </div>
                <div className="flex-1">
                  <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest ml-1">
                    Filter Tanggal
                  </label>
                  <input
                    type="date"
                    value={filterTanggalAbsensi}
                    onChange={(e) => setFilterTanggalAbsensi(e.target.value)}
                    className="w-full bg-zinc-50 border-0 rounded-xl px-4 py-2 font-bold text-sm mt-1"
                  />
                </div>
                <div className="flex items-end gap-2">
                  <select
                    value={pageSize}
                    onChange={(e) => setPageSize(parseInt(e.target.value))}
                    className="bg-zinc-100 text-zinc-500 px-3 py-2 rounded-xl text-[10px] font-black uppercase border-0 focus:ring-0 h-[40px]"
                  >
                    {[10, 20, 50, 100].map((v) => (
                      <option key={v} value={v}>
                        Tampil {v}
                      </option>
                    ))}
                  </select>
                  <button
                    onClick={() => {
                      setFilterClassAbsensi("");
                      setFilterNamaAbsensi("");
                      setFilterTanggalAbsensi("");
                    }}
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
                      {session?.role === "Admin" && (
                        <th className="px-6 py-4 text-center">Hapus</th>
                      )}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-50">
                    {attendance
                      .filter((a) => {
                        const matchesClass = filterClassAbsensi
                          ? a.kelas === filterClassAbsensi
                          : true;
                        const matchesName = filterNamaAbsensi
                          ? a.nama
                              .toLowerCase()
                              .includes(filterNamaAbsensi.toLowerCase())
                          : true;
                        const matchesDate = filterTanggalAbsensi
                          ? a.tanggal === filterTanggalAbsensi
                          : true;
                        return matchesClass && matchesName && matchesDate;
                      })
                      .sort(
                        (a, b) =>
                          b.tanggal.localeCompare(a.tanggal) ||
                          b.jam.localeCompare(a.jam),
                      )
                      .slice(
                        pagination.absensiSiswa,
                        pagination.absensiSiswa + pageSize,
                      )
                      .map((a, i) => (
                        <tr
                          key={a.id}
                          className="hover:bg-gray-50 transition-colors"
                        >
                          <td className="px-6 py-4 text-center font-bold text-gray-400 text-xs">
                            {pagination.absensiSiswa + i + 1}
                          </td>
                          <td className="px-6 py-4 font-medium">{a.tanggal}</td>
                          <td className="px-6 py-4">
                            <p className="font-bold">{a.nama}</p>
                            <p className="text-[10px] text-gray-400 font-mono">
                              {a.nisn}
                            </p>
                          </td>
                          <td className="px-6 py-4 font-bold">{a.kelas}</td>
                          <td className="px-6 py-4 text-center">
                            <span
                              className={`px-2 py-1 rounded-lg text-[10px] font-black border-0 ${
                                a.status === "Hadir"
                                  ? "bg-green-100 text-green-700"
                                  : a.status === "Alfa"
                                    ? "bg-red-100 text-red-700"
                                    : "bg-orange-100 text-orange-700"
                              }`}
                            >
                              {a.status}
                            </span>
                          </td>
                          <td className="px-6 py-4">
                            <p
                              className="text-[10px] font-medium text-gray-500 italic max-w-[150px] truncate"
                              title={a.keterangan || "-"}
                            >
                              {a.keterangan || "-"}
                            </p>
                          </td>
                          <td className="px-6 py-4 text-center">
                            {session?.role === "Admin" || session?.isWali ? (
                              <div className="flex justify-center gap-1">
                                {["Hadir", "Sakit", "Izin", "Alfa"].map(
                                  (stat) => (
                                    <button
                                      key={stat}
                                      title={stat}
                                      onClick={async () => {
                                        toggleLoader(true);
                                        await firestoreService.updateAbsensiStatus(
                                          a.id,
                                          stat,
                                        );
                                        toggleLoader(false);
                                      }}
                                      className={`w-6 h-6 rounded flex items-center justify-center text-[8px] font-bold ${a.status === stat ? "bg-green-800 text-white shadow" : "bg-gray-100 text-gray-400 hover:bg-gray-200"}`}
                                    >
                                      {stat[0]}
                                    </button>
                                  ),
                                )}
                              </div>
                            ) : (
                              <span className="text-gray-300 italic text-[10px]">
                                Read Only
                              </span>
                            )}
                          </td>
                          <td className="px-6 py-4 text-center">
                            {session?.role === "Admin" && (
                              <button
                                onClick={() => {
                                  setConfirmModal({
                                    show: true,
                                    title: "Hapus Log Absensi?",
                                    entityName: `Absensi Siswa: ${a.nama} kelas ${a.kelas} pada tanggal ${a.tanggal}`,
                                    message:
                                      "Rekaman presensi siswa ini akan dihapus dari riwayat hari ini.",
                                    onConfirm: async () => {
                                      toggleLoader(true);
                                      try {
                                        await firestoreService.hapusAbsensi(
                                          a.id,
                                        );
                                        triggerSuccess(
                                          "BERHASIL",
                                          `Log absensi ${a.nama} berhasil dihapus.`,
                                        );
                                      } catch (err) {
                                        alert("Gagal menghapus data.");
                                      } finally {
                                        toggleLoader(false);
                                      }
                                    },
                                  });
                                }}
                                className="text-red-300 hover:text-red-600"
                              >
                                <Trash2 size={16} />
                              </button>
                            )}
                          </td>
                        </tr>
                      ))}
                    {attendance.length === 0 && (
                      <tr>
                        <td
                          colSpan={7}
                          className="text-center py-10 text-gray-400 font-bold uppercase tracking-widest text-xs"
                        >
                          Belum ada rekaman presensi
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
              {attendance.length > pageSize && (
                <div className="mt-4 flex justify-end items-center gap-4 bg-white p-4 rounded-2xl border border-gray-100 shadow-sm text-zinc-900">
                  <span className="text-[10px] font-black text-gray-400 uppercase tracking-widest">
                    Halaman {Math.floor(pagination.absensiSiswa / pageSize) + 1}
                  </span>
                  <div className="flex gap-2">
                    <button
                      disabled={pagination.absensiSiswa === 0}
                      onClick={() =>
                        setPagination({
                          ...pagination,
                          absensiSiswa: pagination.absensiSiswa - pageSize,
                        })
                      }
                      className="p-2 rounded-lg bg-white border border-gray-100 text-gray-400 disabled:opacity-30 hover:bg-gray-50 transition-all"
                    >
                      <ChevronLeft size={16} />
                    </button>
                    <button
                      disabled={
                        pagination.absensiSiswa + pageSize >= attendance.length
                      }
                      onClick={() =>
                        setPagination({
                          ...pagination,
                          absensiSiswa: pagination.absensiSiswa + pageSize,
                        })
                      }
                      className="p-2 rounded-lg bg-white border border-gray-100 text-gray-400 disabled:opacity-30 hover:bg-gray-50 transition-all"
                    >
                      <ChevronRight size={16} />
                    </button>
                  </div>
                </div>
              )}
            </motion.div>
          )}

          {activePanel === "absensi-guru" && (
            <motion.div
              key="absensi-guru"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
            >
              <div className="flex flex-col lg:flex-row gap-3 justify-between items-start lg:items-center mb-6">
                <div className="flex flex-col">
                  <h2 className="text-xl font-bold flex items-center gap-2 text-zinc-900">
                    <Clock size={20} className="text-green-700" /> Log Presensi
                    Mengajar Guru
                  </h2>
                  <p className="text-xs text-zinc-400 font-medium font-semibold">
                    Daftar riwayat presensi tatap muka dan aktivitas mengajar
                  </p>
                </div>
                <div className="flex flex-wrap items-center gap-2 w-full lg:w-auto">
                  {session?.role === "Admin" && (
                    <button
                      onClick={() => {
                        setResetAbsensiTarget("Guru");
                      }}
                      className="bg-red-600 hover:bg-red-700 text-white px-4 py-2.5 rounded-xl text-[10px] font-black tracking-widest uppercase flex items-center gap-2 transition-all shadow-sm"
                    >
                      <Trash2 size={14} /> Reset Absensi Guru
                    </button>
                  )}
                  <select
                    value={pageSize}
                    onChange={(e) => setPageSize(parseInt(e.target.value))}
                    className="bg-zinc-100 text-zinc-600 px-3 py-2.5 rounded-xl text-[10px] font-black uppercase border-0 focus:ring-0 cursor-pointer"
                  >
                    {[10, 20, 50, 100].map((v) => (
                      <option key={v} value={v}>
                        Tampil {v}
                      </option>
                    ))}
                  </select>
                  <select
                    value={filterGuruAbsensiClass}
                    onChange={(e) => {
                      setFilterGuruAbsensiClass(e.target.value);
                      setPagination((prev) => ({ ...prev, absensiGuru: 0 }));
                    }}
                    className="bg-zinc-100 text-zinc-600 px-4 py-2.5 rounded-xl text-xs font-semibold border-0 focus:ring-0 cursor-pointer animate-fade-in"
                  >
                    <option value="">Semua Kelas</option>
                    {classrooms.map((c) => (
                      <option key={c.nama} value={c.nama}>
                        {c.nama}
                      </option>
                    ))}
                  </select>
                  <input
                    type="date"
                    value={filterTanggalAbsensiGuru}
                    onChange={(e) => {
                      setFilterTanggalAbsensiGuru(e.target.value);
                      setPagination((prev) => ({ ...prev, absensiGuru: 0 }));
                    }}
                    className="bg-zinc-100 text-zinc-600 px-4 py-2.5 rounded-xl text-xs font-semibold border-0 focus:ring-0 cursor-pointer animate-fade-in"
                  />
                  {(filterGuruAbsensiClass || filterTanggalAbsensiGuru) && (
                    <button
                      onClick={() => {
                        setFilterGuruAbsensiClass("");
                        setFilterTanggalAbsensiGuru("");
                        setPagination((prev) => ({ ...prev, absensiGuru: 0 }));
                      }}
                      className="bg-zinc-200 hover:bg-zinc-300 text-zinc-700 font-bold text-xs px-4 py-2.5 rounded-xl transition-all"
                    >
                      Clear
                    </button>
                  )}
                </div>
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
                        {session?.role === "Admin" && (
                          <th className="px-6 py-4 text-center">Opsi</th>
                        )}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-50">
                      {filteredTeacherAttendance
                        .slice(
                          pagination.absensiGuru,
                          pagination.absensiGuru + pageSize,
                        )
                        .map((a, i) => (
                          <tr
                            key={a.id}
                            className="hover:bg-gray-50 transition-colors"
                          >
                            <td className="px-6 py-4 text-center font-bold text-gray-400 text-xs">
                              {pagination.absensiGuru + i + 1}
                            </td>
                            <td className="px-6 py-4 text-gray-400 font-mono font-bold">
                              {a.nip}
                            </td>
                            <td className="px-6 py-4 font-bold">{a.nama}</td>
                            <td className="px-6 py-4">
                              <span className="text-zinc-600 font-bold">
                                {a.mapel || "-"}
                              </span>
                            </td>
                            <td className="px-6 py-4">
                              <span className="bg-green-100 text-green-800 px-2 py-0.5 rounded text-[10px] font-bold">
                                {a.kelas}
                              </span>
                            </td>
                            <td className="px-6 py-4 font-medium text-gray-600">
                              {formatIndoDate(a.tanggal)}
                            </td>
                            <td className="px-6 py-4 font-bold">{a.jam}</td>
                            {session?.role === "Admin" && (
                              <td className="px-6 py-4 text-center">
                                <button
                                  onClick={() => {
                                    setConfirmModal({
                                      show: true,
                                      title: "Hapus Absensi Guru?",
                                      entityName: `Absensi Guru: ${a.nama} kelas ${a.kelas} pada tanggal ${a.tanggal}`,
                                      message:
                                        "Rekaman presensi guru ini akan dihapus dari riwayat.",
                                      onConfirm: async () => {
                                        toggleLoader(true);
                                        try {
                                          await firestoreService.hapusAbsensiGuru(
                                            a.id,
                                          );
                                          triggerSuccess(
                                            "BERHASIL",
                                            `Presensi guru ${a.nama} berhasil dihapus.`,
                                          );
                                        } catch (err) {
                                          alert("Gagal menghapus data.");
                                        } finally {
                                          toggleLoader(false);
                                        }
                                      },
                                    });
                                  }}
                                  className="text-red-500 hover:bg-red-50 p-2 rounded-lg transition-all"
                                >
                                  <Trash2 size={16} />
                                </button>
                              </td>
                            )}
                          </tr>
                        ))}
                      {filteredTeacherAttendance.length === 0 && (
                        <tr>
                          <td
                            colSpan={session?.role === "Admin" ? 8 : 7}
                            className="text-center py-8 text-gray-400 font-medium italic"
                          >
                            Belum ada data mengajar hari ini.
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
              {filteredTeacherAttendance.length > pageSize && (
                <div className="mt-4 flex justify-end items-center gap-4 bg-white p-4 rounded-2xl border border-gray-100 shadow-sm text-zinc-900">
                  <span className="text-[10px] font-black text-gray-400 uppercase tracking-widest">
                    Halaman {Math.floor(pagination.absensiGuru / pageSize) + 1}
                  </span>
                  <div className="flex gap-2">
                    <button
                      disabled={pagination.absensiGuru === 0}
                      onClick={() =>
                        setPagination({
                          ...pagination,
                          absensiGuru: pagination.absensiGuru - pageSize,
                        })
                      }
                      className="p-2 rounded-lg bg-white border border-gray-100 text-gray-400 disabled:opacity-30 hover:bg-gray-50 transition-all cursor-pointer"
                    >
                      <ChevronLeft size={16} />
                    </button>
                    <button
                      disabled={
                        pagination.absensiGuru + pageSize >=
                        filteredTeacherAttendance.length
                      }
                      onClick={() =>
                        setPagination({
                          ...pagination,
                          absensiGuru: pagination.absensiGuru + pageSize,
                        })
                      }
                      className="p-2 rounded-lg bg-white border border-gray-100 text-gray-400 disabled:opacity-30 hover:bg-gray-50 transition-all cursor-pointer"
                    >
                      <ChevronRight size={16} />
                    </button>
                  </div>
                </div>
              )}
            </motion.div>
          )}

          {activePanel === "rekap-mapel" &&
            (() => {
              const targetNipForClasses =
                rekapMapelFilter.nip ||
                (session?.role !== "Admin" ? session?.uid : "");
              const allowedClasses = targetNipForClasses
                ? Array.from(
                    new Set(
                      teachingSchedules
                        .filter((s) => s.nip === targetNipForClasses)
                        .map((s) => s.kelas)
                        .filter(Boolean),
                    ),
                  )
                : classrooms.map((c) => c.nama);

              const filteredRecap = (() => {
                if (!rekapMapelFilter.bulan) return [];
                const [yearStr, monthStr] = rekapMapelFilter.bulan.split("-");
                const yearNum = parseInt(yearStr);
                const monthNum = parseInt(monthStr);
                if (isNaN(yearNum) || isNaN(monthNum)) return [];

                const daysInMonth = new Date(yearNum, monthNum, 0).getDate();
                const monthDates: string[] = [];
                for (let d = 1; d <= daysInMonth; d++) {
                  const dayString = d < 10 ? `0${d}` : `${d}`;
                  monthDates.push(`${rekapMapelFilter.bulan}-${dayString}`);
                }

                const targetSelectedNip =
                  rekapMapelFilter.nip ||
                  (session?.role !== "Admin" ? session?.uid : "");
                const matchingSchedules = teachingSchedules.filter((s) => {
                  const matchesNip = targetSelectedNip
                    ? s.nip === targetSelectedNip
                    : true;
                  const matchesMapel = rekapMapelFilter.mapel
                    ? s.mapel === rekapMapelFilter.mapel
                    : true;
                  const matchesKelas = rekapMapelFilter.kelas
                    ? s.kelas === rekapMapelFilter.kelas
                    : true;
                  return matchesNip && matchesMapel && matchesKelas;
                });

                const fullRecapList: any[] = [];
                let virtualIdCounter = 1;

                monthDates.forEach((dateStr) => {
                  const indonesianDayName = getIndonesianDay(dateStr);
                  if (indonesianDayName === "Minggu") return;

                  const holidayMatch = holidays.find(
                    (h) => h.tanggal === dateStr,
                  );
                  const daySett = settings.find(
                    (s) => s.hari === indonesianDayName,
                  );
                  const limitJp = indonesianDayName === "Jumat" ? 6 : 8;
                  const defaultJps = Array.from(
                    { length: limitJp },
                    (_, i) => i + 1,
                  );

                  let activeJps = defaultJps;
                  let reason = "Kegiatan Utama";

                  if (holidayMatch) {
                    activeJps = [];
                    reason = holidayMatch.keterangan;
                  } else if (daySett) {
                    if (!daySett.targetDate || daySett.targetDate === dateStr) {
                      activeJps = Array.isArray(daySett.activeJps)
                        ? daySett.activeJps
                        : defaultJps;
                      reason = daySett.reasonInactive || "Kegiatan Utama";
                    }
                  }

                  const activeSchedulesForDay = matchingSchedules.filter(
                    (s) => s.hari === indonesianDayName,
                  );

                  activeSchedulesForDay.forEach((sch) => {
                    let actualScan = teacherAttendance.find(
                      (ta) =>
                        ta.nip === sch.nip &&
                        ta.tanggal === dateStr &&
                        ta.mapel === sch.mapel &&
                        ta.kelas === sch.kelas,
                    );

                    if (!actualScan) {
                      actualScan = teacherAttendance.find(
                        (ta) =>
                          ta.nip === sch.nip &&
                          ta.tanggal === dateStr &&
                          (ta.kelas === sch.kelas || ta.kelas === "-"),
                      );
                    }

                    const schJps = sch.jps || [];
                    const hasInactiveScheduleJp =
                      schJps.length > 0
                        ? schJps.some((jp) => !activeJps.includes(jp))
                        : activeJps.length < limitJp;

                    let finalStatusText = "Tidak Mengajar";
                    if (actualScan) {
                      if (
                        actualScan.status === "Izin" ||
                        actualScan.status === "Sakit" ||
                        actualScan.status === "Alfa"
                      ) {
                        finalStatusText = actualScan.status;
                      } else {
                        finalStatusText = hasInactiveScheduleJp
                          ? `Mengajar dan ${reason}`
                          : "Mengajar";
                      }
                    } else {
                      if (hasInactiveScheduleJp) {
                        finalStatusText = reason;
                      } else {
                        finalStatusText = "Tidak Mengajar";
                      }
                    }

                    fullRecapList.push({
                      id:
                        actualScan?.id ||
                        `virt-${sch.nip}-${sch.kelas}-${dateStr}-${virtualIdCounter++}`,
                      nip: sch.nip,
                      nama: sch.namaGuru,
                      tanggal: dateStr,
                      mapel: sch.mapel || "-",
                      kelas: sch.kelas,
                      jam: actualScan?.jam || "-",
                      terlambat: actualScan?.terlambat || 0,
                      scanned:
                        !!actualScan &&
                        (!actualScan.status || actualScan.status === "Hadir"),
                      statusKeterangan: finalStatusText,
                      jps: schJps,
                    });
                  });
                });

                return fullRecapList.sort((a, b) =>
                  a.tanggal.localeCompare(b.tanggal),
                );
              })();

              const validRekapMapelOffset =
                pagination.rekapMapel >= filteredRecap.length
                  ? 0
                  : pagination.rekapMapel;
              const slicedRecap = filteredRecap.slice(
                validRekapMapelOffset,
                validRekapMapelOffset + pageSize,
              );

              return (
                <motion.div
                  key="rekap-mapel"
                  initial={{ opacity: 0 }}
                  animate={{ opacity: 1 }}
                >
                  <div className="bg-white rounded-3xl shadow-sm border border-gray-100 p-8 mb-6">
                    <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-6">
                      <div className="flex-1">
                        <h2 className="text-xl font-bold flex items-center gap-2 text-zinc-900">
                          <FileBarChart className="text-green-700" /> Rekap
                          Presensi Mapel
                        </h2>
                        <p className="text-xs text-gray-500 font-medium mt-1 italic">
                          Laporan performa pengajar dan kehadiran siswa di
                          kelas.
                        </p>
                      </div>
                      <div className="flex items-center gap-2">
                        <select
                          value={pageSize}
                          onChange={(e) =>
                            setPageSize(parseInt(e.target.value))
                          }
                          className="bg-zinc-100 text-zinc-600 px-3 py-2 rounded-xl text-[10px] font-black uppercase border-0 focus:ring-0"
                        >
                          {[10, 20, 50, 100].map((v) => (
                            <option key={v} value={v}>
                              Tampil {v}
                            </option>
                          ))}
                        </select>
                      </div>
                      <div className="flex items-center gap-3 w-full md:w-auto">
                        <button
                          onClick={() => {
                            const dataForExcel = filteredRecap.map((a, idx) => {
                              const jpTextCol =
                                a.jps && a.jps.length > 0
                                  ? a.jps
                                      .map((j: number) => `JP ${j}`)
                                      .join(", ")
                                  : "-";
                              return {
                                No: idx + 1,
                                Hari: getIndonesianDay(a.tanggal),
                                Tanggal: a.tanggal,
                                "Nama Guru": a.nama,
                                Mapel: a.mapel || "-",
                                Kelas: a.kelas,
                                JP: jpTextCol,
                                Status: a.statusKeterangan,
                              };
                            });
                            const ws = XLSX.utils.json_to_sheet(dataForExcel);
                            const wb = XLSX.utils.book_new();
                            XLSX.utils.book_append_sheet(wb, ws, "Rekap Mapel");
                            XLSX.writeFile(
                              wb,
                              `Rekap_Mapel_${rekapMapelFilter.bulan}.xlsx`,
                            );
                          }}
                          className="flex-1 md:flex-none flex items-center justify-center gap-2 bg-green-50 text-green-700 px-5 py-3 rounded-2xl text-[10px] font-black uppercase tracking-widest hover:bg-green-100 transition-all border border-green-100 cursor-pointer"
                        >
                          <FileSpreadsheet size={16} /> Excel
                        </button>
                        <button
                          onClick={() => {
                            if (!rekapMapelFilter.kelas || !rekapMapelFilter.mapel) {
                              triggerWarning("FILTER BELUM LENGKAP", "Silakan pilih Filter Kelas dan Filter Mapel terlebih dahulu.");
                              return;
                            }
                            setShowMonthlyMapelAttendance(true);
                          }}
                          className="flex-1 md:flex-none flex items-center justify-center gap-2 bg-blue-50 text-blue-700 px-5 py-3 rounded-2xl text-[10px] font-black uppercase tracking-widest hover:bg-blue-100 transition-all border border-blue-100 cursor-pointer"
                        >
                          <CalendarRange size={16} /> Rekap Absensi
                        </button>
                        <button
                          onClick={() => {
                            const doc = new jsPDF("l", "mm", "a4");
                            doc.setFontSize(16);
                            doc.text(
                              "REKAPITULASI MENGAJAR DAN PRESENSI MAPEL",
                              14,
                              20,
                            );
                            doc.setFontSize(10);
                            doc.text(
                              `Bulan: ${rekapMapelFilter.bulan}`,
                              14,
                              28,
                            );

                            const tableData = filteredRecap.map((a, idx) => {
                              const jpTextCol =
                                a.jps && a.jps.length > 0
                                  ? a.jps
                                      .map((j: number) => `JP ${j}`)
                                      .join(", ")
                                  : "-";
                              return [
                                idx + 1,
                                getIndonesianDay(a.tanggal),
                                a.tanggal,
                                a.nama,
                                a.mapel || "-",
                                a.kelas,
                                jpTextCol,
                                a.statusKeterangan,
                              ];
                            });

                            autoTable(doc, {
                              startY: 35,
                              head: [
                                [
                                  "No",
                                  "Hari",
                                  "Tanggal",
                                  "Nama Guru",
                                  "Mapel",
                                  "Kelas",
                                  "JP",
                                  "Status",
                                ],
                              ],
                              body: tableData,
                              theme: "striped",
                            });

                            doc.save(
                              `Rekap_Mapel_${rekapMapelFilter.bulan}.pdf`,
                            );
                          }}
                          className="flex-1 md:flex-none flex items-center justify-center gap-2 bg-red-50 text-red-700 px-5 py-3 rounded-2xl text-[10px] font-black uppercase tracking-widest hover:bg-red-100 transition-all border border-red-100"
                        >
                          <FileText size={16} /> PDF
                        </button>
                      </div>
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-4 gap-4 mt-8">
                      <div>
                        <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest ml-1">
                          Filter Guru
                        </label>
                        <select
                          disabled={session?.role !== "Admin"}
                          value={
                            session?.role !== "Admin"
                              ? session?.uid
                              : rekapMapelFilter.nip
                          }
                          onChange={(e) =>
                            setRekapMapelFilter({
                              ...rekapMapelFilter,
                              nip: e.target.value,
                              mapel: "",
                              kelas: "",
                            })
                          }
                          className="w-full bg-gray-50 border-0 rounded-xl px-4 py-3 font-bold text-sm mt-1 disabled:opacity-50"
                        >
                          {session?.role === "Admin" ? (
                            <>
                              <option value="">Semua Guru</option>
                              {teachers
                                .filter(
                                  (t) =>
                                    t.role !== "Siswa" &&
                                    t.nip !== "ADMIN001" &&
                                    t.role !== "Admin",
                                )
                                .map((t) => (
                                  <option key={t.nip} value={t.nip}>
                                    {t.nama}
                                  </option>
                                ))}
                            </>
                          ) : (
                            <option value={session?.uid}>
                              {session?.name}
                            </option>
                          )}
                        </select>
                      </div>
                      <div>
                        <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest ml-1">
                          Filter Mapel
                        </label>
                        <select
                          value={rekapMapelFilter.mapel}
                          onChange={(e) =>
                            setRekapMapelFilter({
                              ...rekapMapelFilter,
                              mapel: e.target.value,
                            })
                          }
                          className="w-full bg-gray-50 border-0 rounded-xl px-4 py-3 font-bold text-sm mt-1"
                        >
                          <option value="">Semua Mapel</option>
                          {Array.from(
                            new Set(
                              teachingSchedules
                                .filter((s) => {
                                  const targetNip =
                                    rekapMapelFilter.nip ||
                                    (session?.role !== "Admin"
                                      ? session?.uid
                                      : "");
                                  return targetNip ? s.nip === targetNip : true;
                                })
                                .map((s) => s.mapel)
                                .filter(Boolean),
                            ),
                          ).map((m) => (
                            <option key={m} value={m}>
                              {m}
                            </option>
                          ))}
                        </select>
                      </div>
                      <div>
                        <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest ml-1">
                          Filter Kelas
                        </label>
                        <select
                          value={rekapMapelFilter.kelas}
                          onChange={(e) =>
                            setRekapMapelFilter({
                              ...rekapMapelFilter,
                              kelas: e.target.value,
                            })
                          }
                          className="w-full bg-gray-50 border-0 rounded-xl px-4 py-3 font-bold text-sm mt-1"
                        >
                          <option value="">Semua Kelas</option>
                          {allowedClasses.map((k) => (
                            <option key={k} value={k}>
                              {k}
                            </option>
                          ))}
                        </select>
                      </div>
                      <div>
                        <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest ml-1">
                          Pilih Bulan
                        </label>
                        <input
                          type="month"
                          value={rekapMapelFilter.bulan}
                          onChange={(e) =>
                            setRekapMapelFilter({
                              ...rekapMapelFilter,
                              bulan: e.target.value,
                            })
                          }
                          className="w-full bg-gray-50 border-0 rounded-xl px-4 py-3 font-bold text-sm mt-1"
                        />
                      </div>
                    </div>
                  </div>

                  <div className="bg-white rounded-3xl shadow-sm border border-gray-100 overflow-hidden">
                    <div className="overflow-x-auto">
                      <table className="w-full text-left text-sm min-w-[800px]">
                        <thead className="bg-gray-50 text-gray-400 uppercase text-[10px] font-black tracking-widest border-b border-gray-100">
                          <tr>
                            <th className="px-6 py-4 text-center">No</th>
                            <th className="px-6 py-4 text-nowrap">Hari</th>
                            <th className="px-6 py-4 text-nowrap">Tanggal</th>
                            <th className="px-6 py-4 text-nowrap">Nama Guru</th>
                            <th className="px-6 py-4 text-nowrap">Mapel</th>
                            <th className="px-6 py-4 text-nowrap">Kelas</th>

                            <th className="px-6 py-4 text-center">Detail</th>
                            <th className="px-6 py-4 text-center text-nowrap">
                              JP
                            </th>
                            <th className="px-6 py-4 text-nowrap">Status</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-gray-50">
                          {slicedRecap.map((a, i) => {
                            const jpText =
                              a.jps && a.jps.length > 0
                                ? a.jps.map((j: number) => `JP ${j}`).join(", ")
                                : "-";
                            return (
                              <tr key={a.id} className="hover:bg-gray-50">
                                <td className="px-6 py-4 text-center text-gray-400 font-bold">
                                  {validRekapMapelOffset + i + 1}
                                </td>
                                <td className="px-6 py-4 font-bold text-zinc-600 text-xs">
                                  {getIndonesianDay(a.tanggal)}
                                </td>
                                <td className="px-6 py-4 font-bold text-zinc-600 text-xs text-nowrap">
                                  {a.tanggal}
                                </td>
                                <td className="px-6 py-4 font-bold">
                                  {a.nama}
                                </td>
                                <td className="px-6 py-4 text-green-700 font-black text-xs uppercase">
                                  {a.mapel || "-"}
                                </td>
                                <td className="px-6 py-4 font-bold">
                                  {a.kelas}
                                </td>
                                <td className="px-6 py-4 text-center font-black">
                                  <button
                                    onClick={() =>
                                      setSelectedDetailRecap({
                                        kelas: a.kelas,
                                        tanggal: a.tanggal,
                                        mapel: a.mapel || "-",
                                        namaGuru: a.nama,
                                      })
                                    }
                                    className="text-white bg-green-600 hover:bg-green-700 hover:shadow-md px-3 py-1.5 rounded-xl text-xs font-black tracking-wider uppercase transition-all shadow-sm inline-flex items-center gap-1 cursor-pointer"
                                  >
                                    <Eye size={12} strokeWidth={2.5} /> Detail
                                  </button>
                                </td>
                                <td className="px-6 py-4 text-center text-xs font-black text-green-950 font-mono text-nowrap">
                                  {jpText}
                                </td>
                                <td className="px-6 py-4 italic text-gray-400 font-medium whitespace-nowrap">
                                  {a.statusKeterangan}
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  </div>

                  {filteredRecap.length > pageSize && (
                    <div className="mt-4 flex justify-end items-center gap-4 bg-white p-4 rounded-2xl border border-gray-100 shadow-sm text-zinc-900">
                      <span className="text-[10px] font-black text-gray-400 uppercase tracking-widest">
                        Halaman{" "}
                        {Math.floor(validRekapMapelOffset / pageSize) + 1} /{" "}
                        {Math.ceil(filteredRecap.length / pageSize)}
                      </span>
                      <div className="flex gap-2">
                        <button
                          disabled={validRekapMapelOffset === 0}
                          onClick={() =>
                            setPagination({
                              ...pagination,
                              rekapMapel: validRekapMapelOffset - pageSize,
                            })
                          }
                          className="p-2 rounded-lg bg-white border border-gray-100 text-gray-400 disabled:opacity-30 hover:bg-gray-50 transition-all cursor-pointer"
                        >
                          <ChevronLeft size={16} />
                        </button>
                        <button
                          disabled={
                            validRekapMapelOffset + pageSize >=
                            filteredRecap.length
                          }
                          onClick={() =>
                            setPagination({
                              ...pagination,
                              rekapMapel: validRekapMapelOffset + pageSize,
                            })
                          }
                          className="p-2 rounded-lg bg-white border border-gray-100 text-gray-400 disabled:opacity-30 hover:bg-gray-50 transition-all cursor-pointer"
                        >
                          <ChevronRight size={16} />
                        </button>
                      </div>
                    </div>
                  )}
                </motion.div>
              );
            })()}

          {activePanel === "siswa-data" && (
            <motion.div
              key="siswa-data"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
            >
              <div className="bg-white rounded-3xl shadow-sm border border-gray-100 overflow-hidden max-w-2xl mx-auto">
                <div className="bg-green-800 p-6 text-white flex justify-between items-center">
                  <div>
                    <h2 className="text-lg font-bold">Tentang Data Siswa</h2>
                    <p className="text-green-200 text-xs mt-1">
                      Status data biodata terakhir yang tersimpan di sistem.
                    </p>
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
                          <img
                            src={currentSiswaData.foto}
                            className="w-full h-full object-cover"
                            alt="Profile"
                          />
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
                                    foto: reader.result as string,
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
                    <p className="text-[10px] font-black text-zinc-400 uppercase tracking-widest mt-3">
                      Upload Foto Profil Anda
                    </p>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                    <div>
                      <label className="block text-xs font-bold text-zinc-400 uppercase mb-2 ml-1">
                        NISN (Kunci)
                      </label>
                      <input
                        type="text"
                        readOnly
                        value={currentSiswaData?.nisn || ""}
                        className="w-full bg-zinc-50 border-0 rounded-2xl py-3 px-4 font-bold text-zinc-400 cursor-not-allowed"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-bold text-zinc-400 uppercase mb-2 ml-1">
                        Nama Lengkap
                      </label>
                      <input
                        type="text"
                        value={biodataForm.nama}
                        onChange={(e) =>
                          setBiodataForm({
                            ...biodataForm,
                            nama: e.target.value,
                          })
                        }
                        className="w-full bg-zinc-100 border-0 rounded-2xl py-3 px-4 font-bold focus:ring-2 focus:ring-green-500"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-bold text-zinc-400 uppercase mb-2 ml-1">
                        Jenis Kelamin
                      </label>
                      <select
                        value={biodataForm.jenisKelamin}
                        onChange={(e) =>
                          setBiodataForm({
                            ...biodataForm,
                            jenisKelamin: e.target.value,
                          })
                        }
                        className="w-full bg-zinc-100 border-0 rounded-2xl py-3 px-4 font-bold focus:ring-2 focus:ring-green-500 border-0"
                      >
                        <option value="L">Laki-Laki</option>
                        <option value="P">Perempuan</option>
                      </select>
                    </div>
                    <div>
                      <label className="block text-xs font-bold text-zinc-400 uppercase mb-2 ml-1">
                        Tempat Lahir
                      </label>
                      <input
                        type="text"
                        placeholder="Contoh: Bombana"
                        value={biodataForm.tempat}
                        onChange={(e) =>
                          setBiodataForm({
                            ...biodataForm,
                            tempat: e.target.value,
                          })
                        }
                        className="w-full bg-zinc-100 border-0 rounded-2xl py-3 px-4 font-bold focus:ring-2 focus:ring-green-500"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-bold text-zinc-400 uppercase mb-2 ml-1">
                        Tanggal Lahir
                      </label>
                      <input
                        type="date"
                        value={biodataForm.tgl}
                        onChange={(e) =>
                          setBiodataForm({
                            ...biodataForm,
                            tgl: e.target.value,
                          })
                        }
                        className="w-full bg-zinc-100 border-0 rounded-2xl py-3 px-4 font-bold focus:ring-2 focus:ring-green-500"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-bold text-zinc-400 uppercase mb-2 ml-1">
                        Nama Ayah Kandung
                      </label>
                      <input
                        type="text"
                        value={biodataForm.ayah}
                        onChange={(e) =>
                          setBiodataForm({
                            ...biodataForm,
                            ayah: e.target.value,
                          })
                        }
                        className="w-full bg-zinc-100 border-0 rounded-2xl py-3 px-4 font-bold focus:ring-2 focus:ring-green-500"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-bold text-zinc-400 uppercase mb-2 ml-1">
                        Nama Ibu Kandung
                      </label>
                      <input
                        type="text"
                        value={biodataForm.ibu}
                        onChange={(e) =>
                          setBiodataForm({
                            ...biodataForm,
                            ibu: e.target.value,
                          })
                        }
                        className="w-full bg-zinc-100 border-0 rounded-2xl py-3 px-4 font-bold focus:ring-2 focus:ring-green-500"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-bold text-zinc-400 uppercase mb-2 ml-1">
                        Kelas Aktif (Kunci)
                      </label>
                      <input
                        type="text"
                        readOnly
                        value={currentSiswaData?.kelas || ""}
                        className="w-full bg-zinc-50 border-0 rounded-2xl py-3 px-4 font-bold text-zinc-400 cursor-not-allowed"
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-bold text-zinc-400 uppercase mb-2 ml-1">
                        WhatsApp Ortu (Kunci)
                      </label>
                      <input
                        type="text"
                        readOnly
                        value={currentSiswaData?.hp || ""}
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
                            ...biodataForm,
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

          {activePanel === "profil" && renderProfile()}

          {activePanel === "input-kehadiran-guru" && renderInputKehadiranGuru()}

          {activePanel === "capaian-guru" && renderLaporanCapaianGuru()}

          {activePanel === "roster" && renderRoster()}

          {activePanel === "kbc" && renderKBC()}

          {activePanel === "profil-siswa" && renderProfilSiswa()}

          {activePanel === "rekap" && (
            <motion.div
              key="rekap"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
            >
              <div className="bg-white p-6 rounded-2xl shadow-sm border border-gray-100 mb-6">
                <div className="flex flex-wrap justify-between items-center gap-4 mb-6">
                  <div className="flex gap-2 p-1 bg-gray-100 rounded-xl w-fit">
                    <button
                      onClick={() =>
                        setRekapFilter({ ...rekapFilter, type: "Siswa" })
                      }
                      className={`px-6 py-2 text-xs font-black rounded-lg transition-all ${rekapFilter.type === "Siswa" ? "bg-green-800 text-white shadow" : "text-gray-400 uppercase tracking-widest"}`}
                    >
                      REKAP SISWA
                    </button>
                    <button
                      onClick={() =>
                        setRekapFilter({ ...rekapFilter, type: "Guru" })
                      }
                      className={`px-6 py-2 text-xs font-black rounded-lg transition-all ${rekapFilter.type === "Guru" ? "bg-green-800 text-white shadow" : "text-gray-400 uppercase tracking-widest"}`}
                    >
                      REKAP GURU
                    </button>
                  </div>
                  <div className="flex items-center gap-2">
                    <span className="text-[10px] font-black text-gray-400 uppercase tracking-widest">
                      Tampil:
                    </span>
                    <select
                      value={pageSize}
                      onChange={(e) => {
                        setPageSize(parseInt(e.target.value));
                        setPagination((prev) => ({
                          ...prev,
                          rekapSiswa: 0,
                          rekapGuru: 0,
                        }));
                      }}
                      className="bg-zinc-100 text-zinc-600 px-3 py-2 rounded-xl text-[10px] font-black uppercase border-0 focus:ring-0 cursor-pointer"
                    >
                      {[10, 20, 50, 100].map((v) => (
                        <option key={v} value={v}>
                          Tampil {v}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>
                <div className="grid grid-cols-1 md:grid-cols-4 gap-4 items-end">
                  <div>
                    <label className="block text-[10px] font-black text-gray-400 uppercase mb-1 ml-1">
                      Pilih Bulan
                    </label>
                    <input
                      type="month"
                      value={rekapFilter.bulan}
                      onChange={(e) =>
                        setRekapFilter({
                          ...rekapFilter,
                          bulan: e.target.value,
                        })
                      }
                      className="w-full bg-gray-50 border-0 rounded-xl px-4 py-2 text-sm font-bold"
                    />
                  </div>
                  <div>
                    <label className="block text-[10px] font-black text-gray-400 uppercase mb-1 ml-1">
                      Filter Kelas
                    </label>
                    <select
                      disabled={(() => {
                        const teacher = teachers.find(
                          (t) => t.nip === session?.uid,
                        );
                        const isKamadWakamad =
                          teacher &&
                          ["Kamad", "Wakamad"].includes(teacher.jabatan);
                        return (
                          !isKamadWakamad &&
                          (session?.role === "Guru Wali Kelas" ||
                            session?.isWali)
                        );
                      })()}
                      value={rekapFilter.kelas}
                      onChange={(e) =>
                        setRekapFilter({
                          ...rekapFilter,
                          kelas: e.target.value,
                        })
                      }
                      className="w-full bg-gray-50 border-0 rounded-xl px-4 py-2 text-sm font-bold disabled:opacity-50"
                    >
                      <option value="">Semua Kelas</option>
                      {classrooms.map((k) => (
                        <option key={k.nama} value={k.nama}>
                          {k.nama}
                        </option>
                      ))}
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
                      {rekapFilter.type === "Siswa" ? (
                        <tr>
                          <th className="px-6 py-4 text-center w-12">No.</th>
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
                          <th className="px-6 py-4 text-center w-12">No.</th>
                          <th className="px-6 py-4">Nama Guru</th>
                          <th className="px-6 py-4 text-center">Mapel</th>
                          <th className="px-6 py-4 text-center">
                            Kelas Mengajar
                          </th>
                          <th className="px-6 py-4 text-center">Target Sesi</th>
                          <th className="px-6 py-4 text-center">Aktual</th>
                          <th className="px-6 py-4 text-center">% Kinerja</th>
                        </tr>
                      )}
                    </thead>
                    <tbody className="divide-y divide-gray-50">
                      {rekapFilter.type === "Siswa"
                        ? students
                            .filter((s) =>
                              rekapFilter.kelas
                                ? s.kelas === rekapFilter.kelas
                                : true,
                            )
                            .slice(
                              pagination.rekapSiswa || 0,
                              (pagination.rekapSiswa || 0) + pageSize,
                            )
                            .map((s, i) => {
                              const indexOffset = pagination.rekapSiswa || 0;
                              const monthAttendance =
                                monthAttendanceMap[s.nisn] || [];
                              const hadir = monthAttendance.filter(
                                (a) => a.status === "Hadir",
                              ).length;
                              const izin = monthAttendance.filter(
                                (a) => a.status === "Izin",
                              ).length;
                              const sakit = monthAttendance.filter(
                                (a) => a.status === "Sakit",
                              ).length;
                              const alfa = monthAttendance.filter(
                                (a) => a.status === "Alfa",
                              ).length;
                              const [y, m] = rekapFilter.bulan
                                .split("-")
                                .map(Number);
                              const now = new Date();
                              const limit =
                                now.getFullYear() === y &&
                                now.getMonth() + 1 === m
                                  ? now.getDate()
                                  : undefined;
                              const effCount = getEffectiveDays(
                                y,
                                m - 1,
                                holidays,
                                settings,
                                limit,
                              ).length;
                              const perc =
                                effCount > 0
                                  ? Math.round(
                                      ((hadir + izin + sakit) / effCount) * 100,
                                    )
                                  : 0;

                              return (
                                <tr
                                  key={i}
                                  className="hover:bg-gray-50 transition-colors"
                                >
                                  <td className="px-6 py-4 text-center font-bold text-gray-500 text-xs">
                                    {indexOffset + i + 1}
                                  </td>
                                  <td className="px-6 py-4 font-bold">
                                    {s.nama}
                                  </td>
                                  <td className="px-6 py-4 text-center font-medium">
                                    {s.kelas}
                                  </td>
                                  <td className="px-6 py-4 text-center">
                                    {hadir}
                                  </td>
                                  <td className="px-6 py-4 text-center">
                                    {izin}
                                  </td>
                                  <td className="px-6 py-4 text-center">
                                    {sakit}
                                  </td>
                                  <td className="px-6 py-4 text-center text-red-600 font-bold">
                                    {alfa}
                                  </td>
                                  <td className="px-6 py-4 text-center">
                                    <span
                                      className={`px-2 py-0.5 rounded-full text-[10px] font-black ${perc >= 80 ? "bg-green-100 text-green-700" : "bg-red-100 text-red-700"}`}
                                    >
                                      {perc}%
                                    </span>
                                  </td>
                                </tr>
                              );
                            })
                        : teachers
                            .filter(
                              (t) => t.nip !== "ADMIN001" && t.role !== "Admin",
                            )
                            .filter((t) => {
                              if (rekapFilter.kelas) {
                                return teachingSchedules.some(
                                  (ts) =>
                                    ts.nip === t.nip &&
                                    ts.kelas === rekapFilter.kelas,
                                );
                              }
                              return true;
                            })
                            .slice(
                              pagination.rekapGuru || 0,
                              (pagination.rekapGuru || 0) + pageSize,
                            )
                            .map((t, i) => {
                              const indexOffset = pagination.rekapGuru || 0;
                              const schedules = teachingSchedules.filter(
                                (ts) =>
                                  ts.nip === t.nip &&
                                  (rekapFilter.kelas
                                    ? ts.kelas === rekapFilter.kelas
                                    : true),
                              );
                              const {
                                totalTarget,
                                hadir,
                                sakit,
                                izin,
                                alfa,
                                perc,
                              } = calculateTeacherMonthlyStats(
                                t.nip,
                                rekapFilter.bulan,
                                rekapFilter.kelas,
                                teachingSchedules,
                                holidays,
                                settings,
                                teacherAttendance,
                              );
                              const mapelTaught =
                                Array.from(
                                  new Set(
                                    schedules
                                      .map((sc) => sc.mapel)
                                      .filter(Boolean),
                                  ),
                                ).join(", ") || "-";

                              return (
                                <tr
                                  key={i}
                                  className="hover:bg-gray-50 transition-colors"
                                >
                                  <td className="px-6 py-4 text-center font-bold text-gray-500 text-xs">
                                    {indexOffset + i + 1}
                                  </td>
                                  <td className="px-6 py-4 font-bold">
                                    {t.nama}
                                  </td>
                                  <td className="px-6 py-4 text-center font-semibold text-gray-700 text-xs">
                                    {mapelTaught}
                                  </td>
                                  <td className="px-6 py-4 text-center text-[10px] font-black uppercase text-gray-400">
                                    {rekapFilter.kelas
                                      ? rekapFilter.kelas
                                      : schedules
                                          .map((s) => s.kelas)
                                          .join(", ") || "-"}
                                  </td>
                                  <td className="px-6 py-4 text-center font-bold">
                                    {totalTarget}
                                  </td>
                                  <td className="px-6 py-4 text-center font-bold text-blue-600">
                                    {hadir}
                                  </td>
                                  <td className="px-6 py-4 text-center">
                                    <div className="flex flex-col items-center">
                                      <span
                                        className={`px-2 py-0.5 rounded-full text-[10px] font-black ${perc >= 50 ? "bg-blue-100 text-blue-700" : "bg-red-100 text-red-700"}`}
                                      >
                                        {perc}%
                                      </span>
                                      <span className="text-[9px] text-gray-400 mt-1 font-semibold">
                                        (Hadir: {hadir}, Sakit: {sakit}, Izin:{" "}
                                        {izin}, Alpa: {alfa})
                                      </span>
                                    </div>
                                  </td>
                                </tr>
                              );
                            })}
                    </tbody>
                  </table>
                </div>
                {/* Pagination Controls */}
                {rekapFilter.type === "Siswa"
                  ? (() => {
                      const totalItems = students.filter((s) =>
                        rekapFilter.kelas
                          ? s.kelas === rekapFilter.kelas
                          : true,
                      ).length;
                      if (totalItems > pageSize) {
                        return (
                          <div className="bg-gray-50 px-6 py-4 border-t border-gray-100 flex items-center justify-between">
                            <span className="text-[10px] font-black text-gray-400 uppercase tracking-widest">
                              Halaman{" "}
                              {Math.floor(
                                (pagination.rekapSiswa || 0) / pageSize,
                              ) + 1}
                            </span>
                            <div className="flex gap-2">
                              <button
                                disabled={(pagination.rekapSiswa || 0) === 0}
                                onClick={() =>
                                  setPagination({
                                    ...pagination,
                                    rekapSiswa:
                                      (pagination.rekapSiswa || 0) - pageSize,
                                  })
                                }
                                className="p-2 rounded-lg bg-white border border-gray-100 text-gray-400 disabled:opacity-30 hover:bg-gray-50 transition-all cursor-pointer"
                              >
                                <ChevronLeft size={16} />
                              </button>
                              <button
                                disabled={
                                  (pagination.rekapSiswa || 0) + pageSize >=
                                  totalItems
                                }
                                onClick={() =>
                                  setPagination({
                                    ...pagination,
                                    rekapSiswa:
                                      (pagination.rekapSiswa || 0) + pageSize,
                                  })
                                }
                                className="p-2 rounded-lg bg-white border border-gray-100 text-gray-400 disabled:opacity-30 hover:bg-gray-50 transition-all cursor-pointer"
                              >
                                <ChevronRight size={16} />
                              </button>
                            </div>
                          </div>
                        );
                      }
                      return null;
                    })()
                  : (() => {
                      const totalItems = teachers
                        .filter(
                          (t) => t.nip !== "ADMIN001" && t.role !== "Admin",
                        )
                        .filter((t) => {
                          if (rekapFilter.kelas) {
                            return teachingSchedules.some(
                              (ts) =>
                                ts.nip === t.nip &&
                                ts.kelas === rekapFilter.kelas,
                            );
                          }
                          return true;
                        }).length;
                      if (totalItems > pageSize) {
                        return (
                          <div className="bg-gray-50 px-6 py-4 border-t border-gray-100 flex items-center justify-between">
                            <span className="text-[10px] font-black text-gray-400 uppercase tracking-widest">
                              Halaman{" "}
                              {Math.floor(
                                (pagination.rekapGuru || 0) / pageSize,
                              ) + 1}
                            </span>
                            <div className="flex gap-2">
                              <button
                                disabled={(pagination.rekapGuru || 0) === 0}
                                onClick={() =>
                                  setPagination({
                                    ...pagination,
                                    rekapGuru:
                                      (pagination.rekapGuru || 0) - pageSize,
                                  })
                                }
                                className="p-2 rounded-lg bg-white border border-gray-100 text-gray-400 disabled:opacity-30 hover:bg-gray-50 transition-all cursor-pointer"
                              >
                                <ChevronLeft size={16} />
                              </button>
                              <button
                                disabled={
                                  (pagination.rekapGuru || 0) + pageSize >=
                                  totalItems
                                }
                                onClick={() =>
                                  setPagination({
                                    ...pagination,
                                    rekapGuru:
                                      (pagination.rekapGuru || 0) + pageSize,
                                  })
                                }
                                className="p-2 rounded-lg bg-white border border-gray-100 text-gray-400 disabled:opacity-30 hover:bg-gray-50 transition-all cursor-pointer"
                              >
                                <ChevronRight size={16} />
                              </button>
                            </div>
                          </div>
                        );
                      }
                      return null;
                    })()}
              </div>
            </motion.div>
          )}

          {activePanel === "pengaturan" && (
            <motion.div
              key="pengaturan"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
            >
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
                {/* COLUMN 1 (LEFT): PENGATURAN JAM MASUK, PULANG & JP AKTIF & KALENDER */}
                <div className="space-y-8">
                  {/* Jam Masuk, Pulang & JP Aktif Card */}
                  <div className="bg-white p-6 rounded-3xl shadow-sm border border-gray-100 animate-fadeIn bg-linear-to-b from-white to-zinc-50/50">
                    <div className="flex items-center justify-between mb-6">
                      <div className="flex items-center gap-3">
                        <div className="bg-green-100 p-2.5 rounded-xl text-green-700 border border-green-200">
                          <Clock size={20} className="shrink-0" />
                        </div>
                        <div className="text-left">
                          <h3 className="text-lg font-bold">
                            Jam Masuk, Pulang & JP Aktif
                          </h3>
                          <p className="text-[10px] text-zinc-400 font-bold uppercase tracking-wider">
                            Atur jam masuk, jam pulang, dan tandai JP pelajaran
                            aktif (mempengaruhi kalender)
                          </p>
                        </div>
                      </div>
                      {settings.length === 0 && (
                        <button
                          onClick={async () => {
                            toggleLoader(true);
                            await firestoreService.initializeSettings();
                            toggleLoader(false);
                          }}
                          className="bg-green-600 text-white px-4 py-2 rounded-xl text-[10px] font-black uppercase tracking-widest hover:bg-green-700 font-bold"
                        >
                          Inisialisasi
                        </button>
                      )}
                    </div>

                    <div className="space-y-6">
                      {[...settings]
                        .sort((a, b) => dayOrder[a.hari] - dayOrder[b.hari])
                        .filter((s) => s.hari !== "Minggu")
                        .map((s, idx) => {
                          const today = new Date();
                          const currentDay = today.getDay();
                          const targetDayNum =
                            dayOrder[s.hari] === 7 ? 0 : dayOrder[s.hari];
                          const diff = targetDayNum - currentDay;
                          const dateAtDay = new Date();
                          dateAtDay.setDate(today.getDate() + diff);

                          const dateStr = dateAtDay.toLocaleDateString("en-CA");
                          const holiday = holidays.find(
                            (h) => h.tanggal === dateStr,
                          );

                          const limitJp = s.hari === "Jumat" ? 6 : 8;
                          const defaultJps = Array.from(
                            { length: limitJp },
                            (_, i) => i + 1,
                          );

                          const resActiveJps = holiday
                            ? []
                            : localActiveJps[s.hari] !== undefined
                              ? localActiveJps[s.hari]
                              : s.activeJps && Array.isArray(s.activeJps)
                                ? s.activeJps
                                : defaultJps;
                          const resReasonInactive = holiday
                            ? holiday.keterangan
                            : localReasonInactive[s.hari] !== undefined
                              ? localReasonInactive[s.hari]
                              : s.reasonInactive || "";
                          const resMasuk =
                            localMasuk[s.hari] !== undefined
                              ? localMasuk[s.hari]
                              : s.masuk || "07:15";
                          const resPulang =
                            localPulang[s.hari] !== undefined
                              ? localPulang[s.hari]
                              : s.pulang || "14:00";

                          const isDaySaved = savedDaysStatus[s.hari] || "idle";

                          return (
                            <div
                              key={idx}
                              className={`flex flex-col p-4 rounded-3xl ${holiday ? "bg-red-50/50 border border-red-100" : "bg-gray-50/80 border border-gray-100"} gap-4 transition-all hover:bg-white hover:shadow-sm`}
                            >
                              {/* Header Row */}
                              <div className="flex items-center justify-between">
                                <div className="flex flex-col text-left">
                                  <span
                                    className={`font-black text-sm text-zinc-800 ${holiday ? "text-red-900" : ""}`}
                                  >
                                    Hari {s.hari}
                                  </span>
                                  <span className="text-[8px] text-gray-400 font-black uppercase tracking-widest">
                                    {dateAtDay.toLocaleDateString("id-ID", {
                                      day: "numeric",
                                      month: "short",
                                      year: "numeric",
                                    })}
                                  </span>
                                  {holiday && (
                                    <span className="text-[8px] text-red-650 font-bold uppercase mt-1">
                                      LIBUR: {holiday.keterangan}
                                    </span>
                                  )}
                                </div>

                                <span className="text-[9px] text-green-700 font-black bg-green-50 border border-green-150 px-2.5 py-0.5 rounded-full select-none leading-none">
                                  {resActiveJps.length} dari {limitJp} JP Aktif
                                </span>
                              </div>

                              {/* Jam Masuk & Pulang */}
                              <div className="grid grid-cols-2 gap-4 bg-white p-3 rounded-2xl border border-gray-150/60 shadow-sm">
                                <div className="flex flex-col text-left">
                                  <span className="text-[9px] font-black text-gray-400 uppercase tracking-wider mb-1 pl-1">
                                    Jam Masuk
                                  </span>
                                  <input
                                    type="time"
                                    value={resMasuk}
                                    disabled={!!holiday}
                                    onChange={(e) => {
                                      setLocalMasuk((prev) => ({
                                        ...prev,
                                        [s.hari]: e.target.value,
                                      }));
                                    }}
                                    className="w-full bg-gray-50/50 border border-gray-150 rounded-xl px-3 py-2 text-xs font-bold shadow-sm focus:ring-2 focus:ring-green-500 outline-none disabled:opacity-30"
                                  />
                                </div>
                                <div className="flex flex-col text-left">
                                  <span className="text-[9px] font-black text-gray-400 uppercase tracking-wider mb-1 pl-1">
                                    Jam Pulang
                                  </span>
                                  <input
                                    type="time"
                                    value={resPulang}
                                    disabled={!!holiday}
                                    onChange={(e) => {
                                      setLocalPulang((prev) => ({
                                        ...prev,
                                        [s.hari]: e.target.value,
                                      }));
                                    }}
                                    className="w-full bg-gray-50/50 border border-gray-150 rounded-xl px-3 py-2 text-xs font-bold shadow-sm focus:ring-2 focus:ring-green-500 outline-none disabled:opacity-30"
                                  />
                                </div>
                              </div>

                              {/* GREEN ICON JP ACTIVE SELECTION (As requested) */}
                              <div className="flex flex-col gap-1 text-left bg-white p-3 rounded-2xl border border-gray-150/60 shadow-sm">
                                <span className="text-[9px] font-black text-zinc-400 uppercase tracking-widest leading-none mb-1.5 pl-1">
                                  Pilih Jam Pelajaran (JP) Aktif:
                                </span>
                                <div className="flex flex-wrap gap-1.55">
                                  {Array.from(
                                    { length: limitJp },
                                    (_, i) => i + 1,
                                  ).map((jpNum) => {
                                    const isActive =
                                      resActiveJps.includes(jpNum);
                                    return (
                                      <button
                                        key={jpNum}
                                        type="button"
                                        disabled={!!holiday}
                                        onClick={() => {
                                          let updatedJps: number[];
                                          if (isActive) {
                                            updatedJps = resActiveJps.filter(
                                              (j) => j !== jpNum,
                                            );
                                          } else {
                                            updatedJps = [
                                              ...resActiveJps,
                                              jpNum,
                                            ].sort((a, b) => a - b);
                                          }
                                          setLocalActiveJps((prev) => ({
                                            ...prev,
                                            [s.hari]: updatedJps,
                                          }));
                                        }}
                                        className={`w-7 h-7 flex items-center justify-center rounded-lg text-[10px] font-black transition-all cursor-pointer ${
                                          isActive
                                            ? "bg-green-600 text-white shadow-sm hover:bg-green-700 border border-green-600 animate-fadeIn"
                                            : "bg-zinc-50 border border-zinc-200 text-zinc-400 hover:bg-zinc-100"
                                        } disabled:opacity-40 disabled:cursor-not-allowed`}
                                      >
                                        {jpNum}
                                      </button>
                                    );
                                  })}
                                </div>

                                {/* Input Keterangan bila JP tidak aktif */}
                                {resActiveJps.length < limitJp && (
                                  <div className="space-y-1 text-left mt-3 pt-2 border-t border-dashed border-gray-100">
                                    <label className="block text-[9px] font-black text-orange-650 uppercase tracking-widest pl-1 leading-none">
                                      Keterangan Kegiatan (JP Nonaktif)
                                    </label>
                                    <input
                                      type="text"
                                      placeholder="Misal: Yasinan / Upacara Bendera / Pramuka"
                                      value={resReasonInactive}
                                      disabled={!!holiday}
                                      onChange={(e) => {
                                        setLocalReasonInactive((prev) => ({
                                          ...prev,
                                          [s.hari]: e.target.value,
                                        }));
                                      }}
                                      className="w-full bg-orange-50/50 border border-orange-100/70 rounded-xl px-3 py-1.5 text-xs font-bold text-orange-900 placeholder-orange-300 focus:ring-2 focus:ring-orange-500 outline-none disabled:opacity-75"
                                    />
                                  </div>
                                )}
                              </div>

                              {/* Unified Save Row (Changes to activeJps here will influence the educational calendar) */}
                              <div className="flex flex-wrap gap-2 justify-end pt-3 border-t border-gray-150">
                                <button
                                  type="button"
                                  disabled={
                                    isDaySaved === "saving" || !!holiday
                                  }
                                  onClick={async () => {
                                    setSavedDaysStatus((prev) => ({
                                      ...prev,
                                      [s.hari]: "saving",
                                    }));
                                    try {
                                      // Save jam masuk/pulang, activeJps, reasonInactive, and targetDate. This updates the educational calendar only for that specific date and day.
                                      await firestoreService.savePengaturanHari(
                                        s.hari,
                                        resMasuk,
                                        resPulang,
                                        resActiveJps,
                                        resReasonInactive,
                                        dateStr,
                                        undefined, // keeps original jpTimes intact
                                      );

                                      setSavedDaysStatus((prev) => ({
                                        ...prev,
                                        [s.hari]: "saved",
                                      }));
                                      setLocalActiveJps((prev) => {
                                        const n = { ...prev };
                                        delete n[s.hari];
                                        return n;
                                      });
                                      setLocalReasonInactive((prev) => {
                                        const n = { ...prev };
                                        delete n[s.hari];
                                        return n;
                                      });
                                      setLocalMasuk((prev) => {
                                        const n = { ...prev };
                                        delete n[s.hari];
                                        return n;
                                      });
                                      setLocalPulang((prev) => {
                                        const n = { ...prev };
                                        delete n[s.hari];
                                        return n;
                                      });

                                      setTimeout(() => {
                                        setSavedDaysStatus((prev) => ({
                                          ...prev,
                                          [s.hari]: "idle",
                                        }));
                                      }, 2000);

                                      triggerSuccess(
                                        "BERHASIL",
                                        `Hari ${s.hari} (${dateStr}) berhasil disimpan.`,
                                      );
                                    } catch (err) {
                                      console.error(err);
                                      setSavedDaysStatus((prev) => ({
                                        ...prev,
                                        [s.hari]: "idle",
                                      }));
                                    }
                                  }}
                                  className={`text-[8.5px] font-black uppercase px-3 py-1.5 rounded-lg transition-all cursor-pointer shadow-sm border ${
                                    isDaySaved === "saved"
                                      ? "bg-emerald-600 text-white border-emerald-600"
                                      : isDaySaved === "saving"
                                        ? "bg-zinc-100 text-zinc-400 border-zinc-200 animate-pulse"
                                        : !!holiday
                                          ? "bg-red-100 text-red-400 border-red-200 cursor-not-allowed"
                                          : "bg-green-700 hover:bg-green-800 text-white border-green-700"
                                  }`}
                                >
                                  {isDaySaved === "saved"
                                    ? "Tersimpan ✓"
                                    : isDaySaved === "saving"
                                      ? "Menyimpan..."
                                      : "Simpan JP & Jam"}
                                </button>
                              </div>
                            </div>
                          );
                        })}
                    </div>
                  </div>

                  {/* Kalender Pengaturan Card */}
                  <div className="bg-white p-6 rounded-3xl shadow-sm border border-gray-100 animate-fadeIn">
                    <div className="flex flex-col gap-4 mb-6">
                      <div className="flex items-center gap-3">
                        <div className="bg-orange-100 p-2 rounded-xl text-orange-600">
                          <Calendar size={20} />
                        </div>
                        <h3 className="text-lg font-bold">
                          Kalender Pengaturan
                        </h3>
                      </div>

                      {/* Month Filter */}
                      <div className="flex items-center justify-between gap-2 bg-gray-50 p-2 rounded-2xl border border-gray-100">
                        <button
                          type="button"
                          onClick={() => {
                            const [y, m] = selectedCalendarMonth
                              .split("-")
                              .map(Number);
                            let prevM = m - 1;
                            let prevY = y;
                            if (prevM === 0) {
                              prevM = 12;
                              prevY = y - 1;
                            }
                            const prevStr = `${prevY}-${String(prevM).padStart(2, "0")}`;
                            setSelectedCalendarMonth(prevStr);
                          }}
                          className="px-3 py-1.5 hover:bg-gray-200/70 rounded-xl transition-all text-gray-700 cursor-pointer text-xs font-black uppercase tracking-wider"
                        >
                          &larr; Prev
                        </button>
                        <input
                          type="month"
                          value={selectedCalendarMonth}
                          onChange={(e) =>
                            setSelectedCalendarMonth(e.target.value)
                          }
                          className="bg-white border border-gray-150 rounded-xl px-3 py-1.5 text-xs font-bold shadow-sm text-center text-zinc-800 outline-none focus:ring-2 focus:ring-orange-500 cursor-pointer"
                        />
                        <button
                          type="button"
                          onClick={() => {
                            const [y, m] = selectedCalendarMonth
                              .split("-")
                              .map(Number);
                            let nextM = m + 1;
                            let nextY = y;
                            if (nextM === 13) {
                              nextM = 1;
                              nextY = y + 1;
                            }
                            const nextStr = `${nextY}-${String(nextM).padStart(2, "0")}`;
                            setSelectedCalendarMonth(nextStr);
                          }}
                          className="px-3 py-1.5 hover:bg-gray-200/70 rounded-xl transition-all text-gray-700 cursor-pointer text-xs font-black uppercase tracking-wider"
                        >
                          Next &rarr;
                        </button>
                      </div>
                    </div>

                    {/* legend */}
                    <div className="flex flex-wrap gap-4 mb-5 justify-center text-[9px] font-black uppercase tracking-widest text-zinc-400">
                      <div className="flex items-center gap-1.5 bg-red-50 px-2 py-1 rounded-lg border border-red-100">
                        <span className="w-2.5 h-2.5 rounded-full bg-red-600 shrink-0"></span>
                        <span className="text-red-700">Hari Libur</span>
                      </div>
                      <div className="flex items-center gap-1.5 bg-orange-50 px-2 py-1 rounded-lg border border-orange-100">
                        <span className="w-2.5 h-2.5 rounded-full bg-orange-500 shrink-0"></span>
                        <span className="text-orange-700">Pengaturan JP</span>
                      </div>
                      <div className="flex items-center gap-1.5 bg-gray-50 px-2 py-1 rounded-lg border border-gray-100">
                        <span className="w-2.5 h-2.5 rounded-full bg-gray-300 shrink-0"></span>
                        <span className="text-gray-500">Default Aktif</span>
                      </div>
                    </div>

                    {/* Calendar grid representation */}
                    {(() => {
                      const [cy, cm] = selectedCalendarMonth
                        .split("-")
                        .map(Number);
                      if (isNaN(cy) || !cm) return null;

                      const daysInMonth = new Date(cy, cm, 0).getDate();
                      const firstDayIndex = new Date(cy, cm - 1, 1).getDay(); // Sunday=0, Monday=1...

                      const weekdays = [
                        "Min",
                        "Sen",
                        "Sel",
                        "Rab",
                        "Kam",
                        "Jum",
                        "Sab",
                      ];
                      const monthHolidays = holidays.filter((h) =>
                        h.tanggal.startsWith(selectedCalendarMonth),
                      );
                      const monthSettings = settings.filter(
                        (s) =>
                          s.targetDate &&
                          s.targetDate.startsWith(selectedCalendarMonth),
                      );

                      return (
                        <div className="space-y-6">
                          <div>
                            <div className="grid grid-cols-7 gap-1 text-center mb-2">
                              {weekdays.map((w, idx) => (
                                <span
                                  key={w}
                                  className={`text-[9px] font-black uppercase tracking-wider ${idx === 0 ? "text-red-500" : "text-gray-400"}`}
                                >
                                  {w}
                                </span>
                              ))}
                            </div>

                            <div className="grid grid-cols-7 gap-1 bg-gray-50/50 p-2 rounded-2xl border border-gray-100">
                              {/* Padding cells prior to start of month */}
                              {Array.from({ length: firstDayIndex }).map(
                                (_, padIdx) => (
                                  <div
                                    key={`calendar-pad-${padIdx}`}
                                    className="aspect-square"
                                  ></div>
                                ),
                              )}

                              {/* Days map */}
                              {Array.from({ length: daysInMonth }).map(
                                (_, idx) => {
                                  const dayNum = idx + 1;
                                  const dStr = `${selectedCalendarMonth}-${String(dayNum).padStart(2, "0")}`;
                                  const dayIndex = new Date(
                                    cy,
                                    cm - 1,
                                    dayNum,
                                  ).getDay();
                                  const isSunday = dayIndex === 0;

                                  const isHolidayMatch = holidays.find(
                                    (h) => h.tanggal === dStr,
                                  );
                                  const isSpecialJpMatch = settings.find(
                                    (s) => s.targetDate === dStr,
                                  );

                                  let cellClass =
                                    "bg-white hover:bg-zinc-100 text-zinc-700 border-zinc-200/80";
                                  let descStr = "";

                                  if (isHolidayMatch) {
                                    cellClass =
                                      "bg-red-50 border-red-200 text-red-600 hover:bg-red-100/80";
                                    descStr = `HARI LIBUR: ${isHolidayMatch.keterangan}`;
                                  } else if (isSunday) {
                                    cellClass =
                                      "bg-red-50/30 border-red-100/50 text-red-400/80 hover:bg-red-50";
                                    descStr = "Libur Mingguan (Ahad)";
                                  } else if (isSpecialJpMatch) {
                                    cellClass =
                                      "bg-orange-50 border-orange-200 text-orange-600 hover:bg-orange-100/80";
                                    const limitJp = dayIndex === 5 ? 6 : 8;
                                    const activeCount =
                                      isSpecialJpMatch.activeJps
                                        ? isSpecialJpMatch.activeJps.length
                                        : limitJp;
                                    descStr = `PENGATURAN JP:\n${activeCount} dari ${limitJp} JP Aktif\nKeterangan: ${isSpecialJpMatch.reasonInactive || "Custom"}`;
                                  }

                                  return (
                                    <div
                                      key={`calendar-day-${dayNum}`}
                                      title={descStr || undefined}
                                      className={`relative aspect-square flex flex-col justify-center items-center rounded-xl border text-xs font-black transition-all cursor-help shadow-sm ${cellClass}`}
                                    >
                                      <span>{dayNum}</span>

                                      {/* Indicator Dots */}
                                      <div className="flex gap-0.5 mt-0.5 h-1">
                                        {isHolidayMatch && (
                                          <span className="w-1 h-1 rounded-full bg-red-600"></span>
                                        )}
                                        {isSpecialJpMatch && (
                                          <span className="w-1 h-1 rounded-full bg-orange-600"></span>
                                        )}
                                      </div>

                                      {/* Small trash bin icon for deleting custom JP configurations */}
                                      {isSpecialJpMatch && (
                                        <button
                                          type="button"
                                          onClick={(e) => {
                                            e.stopPropagation();
                                            setConfirmModal({
                                              show: true,
                                              title:
                                                "Hapus Pengaturan JP Khusus?",
                                              message:
                                                "Apakah anda ingin menghapus pengaturan jp pada tanggal yang terpilih?",
                                              onConfirm: async () => {
                                                toggleLoader(true);
                                                try {
                                                  await firestoreService.hapusPengaturanHariKhusus(
                                                    isSpecialJpMatch.hari,
                                                  );

                                                  // Reset local states to default immediately in the UI as well
                                                  setLocalActiveJps((prev) => {
                                                    const n = { ...prev };
                                                    delete n[
                                                      isSpecialJpMatch.hari
                                                    ];
                                                    return n;
                                                  });
                                                  setLocalReasonInactive(
                                                    (prev) => {
                                                      const n = { ...prev };
                                                      delete n[
                                                        isSpecialJpMatch.hari
                                                      ];
                                                      return n;
                                                    },
                                                  );
                                                  setLocalMasuk((prev) => {
                                                    const n = { ...prev };
                                                    delete n[
                                                      isSpecialJpMatch.hari
                                                    ];
                                                    return n;
                                                  });
                                                  setLocalPulang((prev) => {
                                                    const n = { ...prev };
                                                    delete n[
                                                      isSpecialJpMatch.hari
                                                    ];
                                                    return n;
                                                  });

                                                  triggerSuccess(
                                                    "BERHASIL",
                                                    `Pengaturan JP khusus tanggal ${dStr} berhasil dihapus.`,
                                                  );
                                                } catch (err) {
                                                  console.error(
                                                    "Error deleting settings override:",
                                                    err,
                                                  );
                                                  alert(
                                                    "Gagal menghapus pengaturan khusus.",
                                                  );
                                                } finally {
                                                  toggleLoader(false);
                                                }
                                              },
                                            });
                                          }}
                                          className="absolute top-1 right-1 text-orange-600 hover:text-red-700 hover:bg-red-50 p-1.5 rounded-lg transition-all cursor-pointer z-50 bg-white/90 shadow-[0_2px_6px_rgba(0,0,0,0.15)] active:scale-90 flex items-center justify-center border border-orange-100"
                                          title="Hapus pengaturan JP khusus"
                                        >
                                          <Trash2
                                            size={11}
                                            className="text-orange-600 hover:text-red-600 font-bold shrink-0 pointer-events-none"
                                          />
                                        </button>
                                      )}
                                    </div>
                                  );
                                },
                              )}
                            </div>
                          </div>

                          {/* Detail of special events shown in beautiful cards below */}
                          <div className="pt-4 border-t border-gray-100 space-y-3">
                            <h4 className="text-[10px] font-black text-gray-400 uppercase tracking-widest pl-1 mb-1">
                              Agenda & Pengaturan Bulan Ini
                            </h4>

                            {monthHolidays.map((h, hIdx) => (
                              <div
                                key={`month-holiday-${hIdx}`}
                                className="flex items-center gap-3 p-3 bg-red-50/50 border border-red-100 rounded-2xl animate-slideIn"
                              >
                                <span className="w-2.5 h-2.5 rounded-full bg-red-600 shrink-0"></span>
                                <div className="text-left">
                                  <p className="text-xs font-black text-red-900">
                                    {h.tanggal} &mdash; {h.keterangan}
                                  </p>
                                  <p className="text-[9px] font-black uppercase text-red-400 tracking-wider">
                                    Hari Libur Terdaftar
                                  </p>
                                </div>
                              </div>
                            ))}

                            {monthSettings.map((s, sIdx) => {
                              const limitJp = s.hari === "Jumat" ? 6 : 8;
                              const activeCount = s.activeJps
                                ? s.activeJps.length
                                : limitJp;
                              return (
                                <div
                                  key={`month-setting-${sIdx}`}
                                  className="flex items-center gap-3 p-3 bg-orange-50/50 border border-orange-100 rounded-2xl animate-slideIn"
                                >
                                  <span className="w-2.5 h-2.5 rounded-full bg-orange-500 shrink-0"></span>
                                  <div className="text-left">
                                    <p className="text-xs font-black text-orange-900">
                                      {s.targetDate} ({s.hari}) &mdash;{" "}
                                      {s.reasonInactive || "Custom JP"}
                                    </p>
                                    <p className="text-[9px] font-black uppercase text-orange-400 tracking-wider">
                                      Terjadwal: {activeCount} dari {limitJp} JP
                                      Aktif (Masuk: {s.masuk}, Pulang:{" "}
                                      {s.pulang})
                                    </p>
                                  </div>
                                </div>
                              );
                            })}

                            {monthHolidays.length === 0 &&
                              monthSettings.length === 0 && (
                                <p className="text-xs text-gray-400 italic text-center py-4 font-bold uppercase tracking-wider">
                                  Tidak ada agenda khusus pada bulan ini.
                                </p>
                              )}
                          </div>
                        </div>
                      );
                    })()}
                  </div>
                </div>
                {/* COLUMN 2 (RIGHT): PENGATURAN WAKTU JP & HARI LIBUR */}
                <div className="space-y-8">
                  {/* Waktu Jam Pelajaran (JP) Card */}
                  <div className="bg-white p-6 rounded-3xl shadow-sm border border-gray-100 animate-fadeIn bg-linear-to-b from-white to-zinc-50/50">
                    <div className="flex items-center gap-3 mb-6">
                      <div className="bg-green-100 p-2.5 rounded-xl text-green-700 border border-green-200">
                        <Clock size={20} className="shrink-0" />
                      </div>
                      <div className="flex flex-col text-left">
                        <h3 className="text-lg font-bold">
                          Waktu Jam Pelajaran (JP)
                        </h3>
                        <p className="text-[10px] text-zinc-400 font-bold uppercase tracking-wider">
                          Atur jam mulai dan selesai untuk tiap JP pelajaran per
                          hari kerja (khusus hari tersebut)
                        </p>
                      </div>
                    </div>

                    <div className="space-y-6">
                      {[...settings]
                        .sort((a, b) => dayOrder[a.hari] - dayOrder[b.hari])
                        .filter((s) => s.hari !== "Minggu")
                        .map((s, idx) => {
                          const today = new Date();
                          const currentDay = today.getDay();
                          const targetDayNum =
                            dayOrder[s.hari] === 7 ? 0 : dayOrder[s.hari];
                          const diff = targetDayNum - currentDay;
                          const dateAtDay = new Date();
                          dateAtDay.setDate(today.getDate() + diff);

                          const dateStr = dateAtDay.toLocaleDateString("en-CA");
                          const holiday = holidays.find(
                            (h) => h.tanggal === dateStr,
                          );

                          const limitJp = s.hari === "Jumat" ? 6 : 8;
                          const defaultJps = Array.from(
                            { length: limitJp },
                            (_, i) => i + 1,
                          );

                          const resMasuk =
                            localMasuk[s.hari] !== undefined
                              ? localMasuk[s.hari]
                              : s.masuk || "07:15";
                          const resPulang =
                            localPulang[s.hari] !== undefined
                              ? localPulang[s.hari]
                              : s.pulang || "14:00";

                          const isDaySaved = savedDaysStatus[s.hari] || "idle";

                          return (
                            <div
                              key={idx}
                              className={`flex flex-col p-4 rounded-3xl ${holiday ? "bg-red-50/50 border border-red-100" : "bg-gray-50/80 border border-gray-100"} gap-3 transition-all hover:bg-white hover:shadow-sm`}
                            >
                              <div className="flex items-center justify-between shadow-xs p-3 rounded-2xl bg-white border border-gray-150/50">
                                <div className="flex flex-col text-left">
                                  <span
                                    className={`font-black text-sm text-zinc-800 ${holiday ? "text-red-900" : ""}`}
                                  >
                                    JP Hari {s.hari}
                                  </span>
                                  <span className="text-[8px] text-gray-400 font-black uppercase tracking-widest">
                                    {dateAtDay.toLocaleDateString("id-ID", {
                                      day: "numeric",
                                      month: "short",
                                      year: "numeric",
                                    })}
                                  </span>
                                  {holiday && (
                                    <span className="text-[8px] text-red-650 font-bold uppercase mt-1">
                                      LIBUR: {holiday.keterangan}
                                    </span>
                                  )}
                                </div>
                              </div>

                              {/* JP Time Settings Editor */}
                              <div className="bg-white p-3 rounded-2xl border border-gray-150/65 shadow-sm flex flex-col gap-2">
                                <button
                                  type="button"
                                  onClick={() =>
                                    setExpandedJpDays((prev) => ({
                                      ...prev,
                                      [s.hari]: !prev[s.hari],
                                    }))
                                  }
                                  className="flex items-center justify-between w-full text-zinc-555 hover:text-zinc-850 transition-colors py-1 cursor-pointer outline-none"
                                >
                                  <span className="text-[9px] font-black uppercase tracking-widest flex items-center gap-1.5 text-zinc-500">
                                    <Clock
                                      size={12}
                                      className="text-green-600 shrink-0"
                                    />
                                    Atur Waktu Tiap JP:
                                  </span>
                                  <span className="text-[9px] font-black tracking-wider text-green-700 bg-green-50 border border-green-100 px-2.5 py-0.5 rounded-full select-none">
                                    {expandedJpDays[s.hari]
                                      ? "Tutup ▲"
                                      : "Buka ▼"}
                                  </span>
                                </button>

                                {expandedJpDays[s.hari] && (
                                  <div className="mt-1 space-y-3 bg-zinc-50 p-3 rounded-2xl border border-zinc-150 animate-fadeIn">
                                    <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-2">
                                      <span className="text-[8px] font-black uppercase tracking-widest text-zinc-400">
                                        Jam Mulai & Selesai JP
                                      </span>
                                      <button
                                        type="button"
                                        onClick={() => {
                                          const standardTimes: any = {};
                                          let startHour = 7;
                                          let startMin = 15;
                                          for (let i = 1; i <= limitJp; i++) {
                                            if (i === 5) {
                                              startMin += 20; // 20 min break
                                              if (startMin >= 60) {
                                                startHour += Math.floor(
                                                  startMin / 60,
                                                );
                                                startMin = startMin % 60;
                                              }
                                            }
                                            const sHourStr = String(
                                              startHour,
                                            ).padStart(2, "0");
                                            const sMinStr = String(
                                              startMin,
                                            ).padStart(2, "0");

                                            startMin += 40; // Each JP is 40 minutes
                                            if (startMin >= 60) {
                                              startHour += Math.floor(
                                                startMin / 60,
                                              );
                                              startMin = startMin % 60;
                                            }
                                            const eHourStr = String(
                                              startHour,
                                            ).padStart(2, "0");
                                            const eMinStr = String(
                                              startMin,
                                            ).padStart(2, "0");

                                            standardTimes[i] = {
                                              start: `${sHourStr}:${sMinStr}`,
                                              end: `${eHourStr}:${eMinStr}`,
                                            };
                                          }
                                          setLocalJpTimes((prev) => ({
                                            ...prev,
                                            [s.hari]: standardTimes,
                                          }));
                                        }}
                                        className="text-[8px] font-black text-green-750 uppercase bg-green-100 border border-green-250 hover:bg-green-200 px-2.5 py-1 rounded-lg transition-all cursor-pointer"
                                      >
                                        Auto-Fill Waktu Standar JP
                                      </button>
                                    </div>

                                    <div className="grid grid-cols-2 gap-2">
                                      {Array.from(
                                        { length: limitJp },
                                        (_, i) => i + 1,
                                      ).map((jpNumber) => {
                                        const originalJpData = s.jpTimes?.[
                                          jpNumber
                                        ] || { start: "", end: "" };
                                        const localJpData =
                                          localJpTimes[s.hari]?.[jpNumber];
                                        const jpStartVal =
                                          localJpData?.start !== undefined
                                            ? localJpData.start
                                            : originalJpData.start;
                                        const jpEndVal =
                                          localJpData?.end !== undefined
                                            ? localJpData.end
                                            : originalJpData.end;

                                        return (
                                          <div
                                            key={jpNumber}
                                            className="flex flex-col sm:flex-row items-start sm:items-center justify-between p-2 bg-white rounded-xl border border-gray-150/60 hover:shadow-xs transition-all gap-1.5 animate-fadeIn"
                                          >
                                            <span className="text-[9px] font-black text-zinc-550">
                                              JP {jpNumber}
                                            </span>
                                            <div className="flex items-center gap-1">
                                              <input
                                                type="time"
                                                value={jpStartVal}
                                                onChange={(e) => {
                                                  const val = e.target.value;
                                                  setLocalJpTimes((prev) => {
                                                    const cur = prev[s.hari]
                                                      ? { ...prev[s.hari] }
                                                      : {
                                                          ...(s.jpTimes || {}),
                                                        };
                                                    cur[jpNumber] = {
                                                      ...(cur[jpNumber] || {
                                                        start: "",
                                                        end: "",
                                                      }),
                                                      start: val,
                                                    };
                                                    return {
                                                      ...prev,
                                                      [s.hari]: cur,
                                                    };
                                                  });
                                                }}
                                                className="w-16 text-center bg-gray-50 border border-gray-150 rounded px-1.5 py-0.5 text-[10px] font-bold focus:ring-1 focus:ring-green-500 outline-none"
                                              />
                                              <span className="text-[10px] text-gray-400 font-extrabold">
                                                &mdash;
                                              </span>
                                              <input
                                                type="time"
                                                value={jpEndVal}
                                                onChange={(e) => {
                                                  const val = e.target.value;
                                                  setLocalJpTimes((prev) => {
                                                    const cur = prev[s.hari]
                                                      ? { ...prev[s.hari] }
                                                      : {
                                                          ...(s.jpTimes || {}),
                                                        };
                                                    cur[jpNumber] = {
                                                      ...(cur[jpNumber] || {
                                                        start: "",
                                                        end: "",
                                                      }),
                                                      end: val,
                                                    };
                                                    return {
                                                      ...prev,
                                                      [s.hari]: cur,
                                                    };
                                                  });
                                                }}
                                                className="w-16 text-center bg-gray-50 border border-gray-150 rounded px-1.5 py-0.5 text-[10px] font-bold focus:ring-1 focus:ring-green-500 outline-none"
                                              />
                                            </div>
                                          </div>
                                        );
                                      })}
                                    </div>
                                  </div>
                                )}
                              </div>

                              {/* Direct Day Saving for JP time edits */}
                              <div className="flex justify-end pt-2 border-t border-gray-100">
                                <button
                                  type="button"
                                  disabled={
                                    isDaySaved === "saving" || !!holiday
                                  }
                                  onClick={async () => {
                                    setSavedDaysStatus((prev) => ({
                                      ...prev,
                                      [s.hari]: "saving",
                                    }));
                                    try {
                                      const mergedJpTimes = {
                                        ...(s.jpTimes || {}),
                                        ...(localJpTimes[s.hari] || {}),
                                      };

                                      await firestoreService.savePengaturanHari(
                                        s.hari,
                                        resMasuk,
                                        resPulang,
                                        undefined,
                                        undefined,
                                        undefined,
                                        mergedJpTimes,
                                      );

                                      setSavedDaysStatus((prev) => ({
                                        ...prev,
                                        [s.hari]: "saved",
                                      }));
                                      setLocalJpTimes((prev) => {
                                        const n = { ...prev };
                                        delete n[s.hari];
                                        return n;
                                      });

                                      setTimeout(() => {
                                        setSavedDaysStatus((prev) => ({
                                          ...prev,
                                          [s.hari]: "idle",
                                        }));
                                      }, 2000);

                                      triggerSuccess(
                                        "BERHASIL",
                                        `Waktu Jam Pelajaran Hari ${s.hari} berhasil disimpan.`,
                                      );
                                    } catch (err) {
                                      console.error(err);
                                      setSavedDaysStatus((prev) => ({
                                        ...prev,
                                        [s.hari]: "idle",
                                      }));
                                    }
                                  }}
                                  className={`text-[8.5px] font-black uppercase px-3 py-1.5 rounded-lg transition-all cursor-pointer shadow-sm border ${
                                    isDaySaved === "saved"
                                      ? "bg-emerald-600 text-white border-emerald-600"
                                      : isDaySaved === "saving"
                                        ? "bg-zinc-100 text-zinc-400 border-zinc-200 animate-pulse"
                                        : !!holiday
                                          ? "bg-red-100 text-red-400 border-red-200 cursor-not-allowed"
                                          : "bg-green-700 hover:bg-green-800 text-white border-green-700"
                                  }`}
                                >
                                  {isDaySaved === "saved"
                                    ? "Tersimpan ✓"
                                    : isDaySaved === "saving"
                                      ? "Menyimpan..."
                                      : "Simpan Waktu Jam Pelajaran"}
                                </button>
                              </div>
                            </div>
                          );
                        })}
                    </div>
                  </div>

                  {/* Hari Libur Card */}
                  <div className="bg-white p-6 rounded-3xl shadow-sm border border-gray-100 animate-fadeIn">
                    <div className="flex items-center gap-3 mb-6">
                      <div className="bg-red-100 p-2 rounded-xl text-red-650 border border-red-205">
                        <Calendar size={20} />
                      </div>
                      <h3 className="text-lg font-bold">Hari Libur</h3>
                    </div>

                    <div className="bg-gray-50/85 p-5 rounded-2xl border-2 border-dashed border-gray-200 mb-6 text-left">
                      <h4 className="text-[10px] font-black text-gray-400 uppercase tracking-widest pl-1 mb-2">
                        Form Tambah Hari Libur
                      </h4>
                      <div className="space-y-3">
                        <input
                          type="date"
                          value={newLibur.tanggal}
                          onChange={(e) =>
                            setNewLibur({
                              ...newLibur,
                              tanggal: e.target.value,
                            })
                          }
                          className="w-full bg-white border border-gray-150 rounded-xl px-4 py-2.5 text-xs font-bold outline-none"
                        />
                        <input
                          type="text"
                          placeholder="Keterangan Hari Libur..."
                          value={newLibur.keterangan}
                          onChange={(e) =>
                            setNewLibur({
                              ...newLibur,
                              keterangan: e.target.value,
                            })
                          }
                          className="w-full bg-white border border-gray-150 rounded-xl px-4 py-2.5 text-xs font-bold outline-none"
                        />
                        <button
                          type="button"
                          onClick={async () => {
                            if (!newLibur.tanggal || !newLibur.keterangan)
                              return;
                            toggleLoader(true);
                            try {
                              await firestoreService.saveLiburAgenda(
                                newLibur.tanggal,
                                newLibur.keterangan,
                              );
                              setNewLibur({ tanggal: "", keterangan: "" });
                              triggerSuccess(
                                "BERHASIL",
                                "Hari libur ditambahkan.",
                              );
                            } catch (e) {
                              alert("Gagal menyimpan.");
                            } finally {
                              toggleLoader(false);
                            }
                          }}
                          className="w-full bg-red-600 text-white rounded-xl py-2.5 font-bold hover:bg-red-500 transition-all text-xs cursor-pointer text-center select-none shadow-sm font-bold uppercase tracking-wider"
                        >
                          Tambah Hari Libur
                        </button>
                      </div>
                    </div>

                    <div className="space-y-3">
                      <h4 className="text-[10px] font-black text-gray-400 uppercase tracking-widest pl-1 text-left block">
                        Daftar Hari Libur Terdaftar
                      </h4>
                      {holidays.length === 0 ? (
                        <p className="text-xs text-gray-400 italic font-bold p-2 text-center bg-gray-50/50 rounded-xl">
                          Belum ada hari libur khusus yang ditambahkan.
                        </p>
                      ) : (
                        holidays.map((h, idx) => (
                          <div
                            key={idx}
                            className="flex items-center justify-between p-4 bg-red-50/50 rounded-2xl border border-red-100 transition-all hover:bg-red-50"
                          >
                            <div className="text-left">
                              <p className="font-black text-red-900 text-sm">
                                {h.keterangan}
                              </p>
                              <p className="text-xs text-red-500/80 font-bold font-mono">
                                {h.tanggal}
                              </p>
                            </div>
                            <button
                              type="button"
                              onClick={async () => {
                                if (
                                  confirm(`Hapus hari libur ${h.keterangan}?`)
                                ) {
                                  toggleLoader(true);
                                  try {
                                    await firestoreService.hapusLiburAgenda(
                                      h.tanggal,
                                    );
                                    triggerSuccess(
                                      "BERHASIL",
                                      "Hari libur berhasil dihapus.",
                                    );
                                  } catch (e) {
                                    alert("Gagal menghapus.");
                                  } finally {
                                    toggleLoader(false);
                                  }
                                }
                              }}
                              className="text-red-400 hover:text-red-600 p-2 cursor-pointer transition-colors"
                            >
                              <Trash2 size={16} />
                            </button>
                          </div>
                        ))
                      )}
                    </div>
                  </div>
                </div>
                {/* Logo & Identity Configuration */}{" "}
                {/* Logo & Identity Configuration */}
                <div className="bg-white p-6 rounded-3xl shadow-sm border border-gray-100 lg:col-span-2 space-y-8">
                  <div className="flex items-center gap-3 mb-2">
                    <div className="bg-blue-100 p-2 rounded-xl text-blue-600">
                      <LayoutGrid size={20} />
                    </div>
                    <h3 className="text-lg font-bold">
                      Identitas & Desain Aplikasi
                    </h3>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-8 items-start">
                    <div className="space-y-6">
                      {/* Logo Section */}
                      <div>
                        <label className="block text-[10px] font-black text-gray-400 uppercase tracking-widest mb-2 ml-1">
                          Logo Aplikasi (Sekolah)
                        </label>

                        <div className="flex flex-col gap-3 mb-4">
                          <input
                            type="text"
                            placeholder="Link Logo (URL)..."
                            defaultValue={appConfig.logoUrl || ""}
                            onBlur={async (e) => {
                              toggleLoader(true);
                              await setDoc(
                                doc(db, "appConfig", "general"),
                                { logoUrl: e.target.value },
                                { merge: true },
                              );
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
                                console.error(
                                  "[DEBUG] Logo load error (Settings Preview):",
                                  e.currentTarget.src,
                                );
                                if (
                                  e.currentTarget.src !== fallbackLogo &&
                                  !e.currentTarget.src.includes("/logo.png")
                                ) {
                                  e.currentTarget.src = fallbackLogo;
                                } else {
                                  e.currentTarget.style.opacity = "0";
                                  const parent = e.currentTarget.parentElement;
                                  if (parent) {
                                    const icon =
                                      parent.querySelector(".fallback-icon");
                                    if (icon)
                                      (icon as HTMLElement).style.opacity = "1";
                                  }
                                }
                              }}
                            />
                            <School
                              size={36}
                              className="text-green-700 absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 opacity-0 fallback-icon transition-opacity"
                            />
                          </div>
                          <div className="flex-1 space-y-3">
                            <p className="text-[10px] text-gray-400 font-medium">
                              Gunakan gambar transparan (PNG) untuk hasil
                              terbaik.
                            </p>
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
                                          throw new Error(
                                            "Ukuran gambar terlalu besar. Gunakan gambar di bawah 750KB.",
                                          );
                                        }
                                        await setDoc(
                                          doc(db, "appConfig", "general"),
                                          {
                                            logoUrl: result,
                                          },
                                          { merge: true },
                                        );
                                        console.log("Logo updated on server");
                                        alert(
                                          "Logo aplikasi berhasil diperbarui!",
                                        );
                                      } catch (err: any) {
                                        console.error(
                                          "Logo upload error:",
                                          err,
                                        );
                                        alert(
                                          "Gagal memperbarui logo: " +
                                            (err.message || "Unknown error"),
                                        );
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
                                    title: "Hapus Logo?",
                                    message:
                                      "Hapus logo kustom dan kembali ke default?",
                                    onConfirm: async () => {
                                      toggleLoader(true);
                                      await setDoc(
                                        doc(db, "appConfig", "general"),
                                        { logoUrl: null },
                                        { merge: true },
                                      );
                                      toggleLoader(false);
                                    },
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
                      <div className="bg-gray-50 p-6 rounded-[2rem] border border-gray-100 shadow-sm">
                        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-6">
                          <div>
                            <label className="block text-[11px] font-black text-gray-500 uppercase tracking-widest ml-1">
                              Informasi Mata Pelajaran Madrasah
                            </label>
                            <p className="text-[11px] text-gray-400 mt-1 ml-1">
                              Total mata pelajaran aktif:{" "}
                              <span className="font-extrabold text-zinc-700">
                                {activeSubjects.length}
                              </span>
                            </p>
                          </div>

                          <div className="flex items-center gap-2">
                            <button
                              onClick={async () => {
                                setConfirmModal({
                                  show: true,
                                  title: "Reset Mapel?",
                                  message:
                                    "Apakah Anda yakin ingin menyetel ulang kustomisasi mata pelajaran ke daftar default madrasah?",
                                  onConfirm: async () => {
                                    toggleLoader(true);
                                    try {
                                      await setDoc(
                                        doc(db, "appConfig", "general"),
                                        { subjects: [...DEFAULT_SUBJECTS] },
                                        { merge: true },
                                      );
                                    } catch (err) {
                                      console.error(err);
                                    } finally {
                                      toggleLoader(false);
                                    }
                                  },
                                });
                              }}
                              className="bg-zinc-100 text-zinc-600 hover:bg-zinc-200 px-3 py-1.5 rounded-xl text-[9px] font-black uppercase tracking-wider transition-all flex items-center gap-1 border border-zinc-200"
                              title="Reset daftar mata pelajaran ke daftar standard madrasah"
                            >
                              <RefreshCcw size={10} /> Reset Default
                            </button>

                            <button
                              onClick={async () => {
                                setConfirmModal({
                                  show: true,
                                  title: "Hapus Semua?",
                                  message:
                                    "Apakah Anda yakin ingin menghapus semua mata pelajaran?",
                                  onConfirm: async () => {
                                    toggleLoader(true);
                                    try {
                                      await setDoc(
                                        doc(db, "appConfig", "general"),
                                        { subjects: [] },
                                        { merge: true },
                                      );
                                    } catch (err) {
                                      console.error(err);
                                    } finally {
                                      toggleLoader(false);
                                    }
                                  },
                                });
                              }}
                              className="bg-red-50 hover:bg-red-100 text-red-600 px-3 py-1.5 rounded-xl text-[9px] font-black uppercase tracking-wider transition-all border border-red-100"
                            >
                              Hapus Semua
                            </button>
                          </div>
                        </div>

                        {/* Input Form Tambah Mapel */}
                        <form
                          onSubmit={async (e) => {
                            e.preventDefault();
                            const trimmed = newSubject.trim();
                            if (!trimmed) return;

                            if (
                              activeSubjects.some(
                                (s) =>
                                  s.toLowerCase() === trimmed.toLowerCase(),
                              )
                            ) {
                              alert(
                                `Mata pelajaran "${trimmed}" sudah terdaftar!`,
                              );
                              return;
                            }

                            toggleLoader(true);
                            try {
                              const updated = [...activeSubjects, trimmed];
                              await setDoc(
                                doc(db, "appConfig", "general"),
                                { subjects: updated },
                                { merge: true },
                              );
                              setNewSubject("");
                            } catch (err) {
                              console.error(err);
                              alert("Gagal menambahkan mata pelajaran.");
                            } finally {
                              toggleLoader(false);
                            }
                          }}
                          className="flex gap-2 mb-6 bg-white p-2 rounded-2xl border border-gray-150 shadow-inner"
                        >
                          <div className="relative flex-1 flex items-center pl-3">
                            <BookOpen
                              size={16}
                              className="text-gray-400 mr-2"
                            />
                            <input
                              type="text"
                              value={newSubject}
                              onChange={(e) => setNewSubject(e.target.value)}
                              placeholder="Masukkan nama mata pelajaran baru..."
                              className="w-full border-0 p-1 text-sm font-bold text-gray-800 placeholder-gray-400 focus:outline-none focus:ring-0"
                            />
                          </div>
                          <button
                            type="submit"
                            className="bg-green-600 hover:bg-green-700 text-white rounded-xl px-4 py-2 text-xs font-black uppercase tracking-widest transition-all flex items-center gap-1 shrink-0"
                          >
                            <Plus size={14} strokeWidth={3} /> Tambah
                          </button>
                        </form>

                        <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 gap-2.5">
                          {activeSubjects.map((subject) => {
                            return (
                              <div
                                key={subject}
                                className="group flex items-center justify-between p-3.5 rounded-xl border border-gray-150 bg-white hover:border-green-200 hover:bg-green-50/20 transition-all shadow-sm animate-fade-in"
                              >
                                <div className="flex items-center gap-2 overflow-hidden">
                                  <div className="w-1.5 h-1.5 rounded-full bg-green-500 shrink-0" />
                                  <span className="text-xs font-extrabold text-zinc-700 truncate">
                                    {subject}
                                  </span>
                                </div>
                                <button
                                  onClick={async () => {
                                    setConfirmModal({
                                      show: true,
                                      title: "Hapus Mapel?",
                                      message: `Apakah Anda yakin ingin menghapus mata pelajaran "${subject}"? Perubahan ini akan segera disinkronisasikan ke seluruh jadwal mengajar.`,
                                      onConfirm: async () => {
                                        toggleLoader(true);
                                        try {
                                          const updated = activeSubjects.filter(
                                            (s) => s !== subject,
                                          );
                                          await setDoc(
                                            doc(db, "appConfig", "general"),
                                            { subjects: updated },
                                            { merge: true },
                                          );
                                        } catch (err) {
                                          console.error(err);
                                        } finally {
                                          toggleLoader(false);
                                        }
                                      },
                                    });
                                  }}
                                  className="text-gray-300 hover:text-red-600 p-1 rounded-lg hover:bg-red-50 transition-all shrink-0"
                                  title={`Hapus ${subject}`}
                                >
                                  <Trash2 size={13} />
                                </button>
                              </div>
                            );
                          })}

                          {activeSubjects.length === 0 && (
                            <div className="col-span-full py-10 text-center bg-white border border-dashed border-gray-200 rounded-2xl">
                              <AlertCircle
                                className="mx-auto text-gray-300 mb-2"
                                size={24}
                              />
                              <p className="text-xs font-extrabold text-gray-400 uppercase tracking-wider">
                                Belum Ada Mata Pelajaran
                              </p>
                              <p className="text-[10px] text-gray-400 mt-1">
                                Silakan ketik nama mapel di atas atau klik
                                'Reset Default' untuk memulihkan daftar bawaan.
                              </p>
                            </div>
                          )}
                        </div>
                        <p className="mt-4 text-[9px] text-gray-400 italic">
                          Mata pelajaran yang terdaftar di atas akan otomatis
                          disinkronkan ke dropdown jadwal mengajar madrasah.
                        </p>
                      </div>
                    </div>

                    <div className="bg-zinc-950 p-6 rounded-[2rem] text-white">
                      <h4 className="text-xs font-black uppercase tracking-[0.2em] text-green-500 mb-4">
                        Informasi Sistem
                      </h4>
                      <div className="space-y-4">
                        <div className="p-4 bg-white/5 rounded-2xl border border-white/5">
                          <p className="text-[9px] font-black text-gray-500 uppercase mb-1">
                            Nama Aplikasi
                          </p>
                          <p className="font-bold text-sm">
                            SIGAP MTsN 2 BOMBANA
                          </p>
                        </div>
                        <div className="p-4 bg-white/5 rounded-2xl border border-white/5">
                          <p className="text-[9px] font-black text-gray-500 uppercase mb-1">
                            Versi
                          </p>
                          <p className="font-bold text-sm">
                            Production v1.2.0-PRO
                          </p>
                        </div>
                        <div className="p-4 bg-white/5 rounded-2xl border border-white/5">
                          <p className="text-[9px] font-black text-gray-500 uppercase mb-1">
                            URL Logo Aplikasi
                          </p>
                          <input
                            type="text"
                            value={appLogoInput}
                            onChange={(e) => setAppLogoInput(e.target.value)}
                            placeholder="https://example.com/logo.png"
                            className="w-full bg-white/10 border-0 rounded-xl py-2 px-3 text-xs font-bold text-white mt-1 focus:ring-1 focus:ring-green-500"
                          />
                        </div>
                        <button
                          onClick={() => {
                            if (!appLogoInput) return;
                            toggleLoader(true);
                            firestoreService
                              .saveAppConfig({ logoUrl: appLogoInput })
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

          {selectedStudentCard &&
            (session?.role === "Siswa" || session?.role === "Admin") && (
              <div
                className="fixed inset-0 z-[200] flex items-center justify-center p-2 sm:p-4 overflow-y-auto"
                style={{
                  backgroundColor: "rgba(0, 0, 0, 0.6)",
                  backdropFilter: "blur(4px)",
                }}
              >
                <div
                  className="absolute inset-0"
                  onClick={() => setSelectedStudentCard(null)}
                />
                <motion.div
                  initial={{ scale: 0.9, opacity: 0 }}
                  animate={{ scale: 1, opacity: 1 }}
                  className="relative p-4 md:p-8 rounded-[2rem] md:rounded-[3rem] shadow-2xl flex flex-col items-center border w-fit max-w-[95vw] sm:max-w-none max-h-[95vh] overflow-y-auto scale-[0.75] sm:scale-[0.85] md:scale-90 lg:scale-100 my-auto origin-top"
                  style={{
                    backgroundColor: "rgba(255, 255, 255, 0.1)",
                    backdropFilter: "blur(16px)",
                    borderColor: "rgba(255, 255, 255, 0.2)",
                  }}
                >
                  <div
                    id="print-student-card"
                    className="flex flex-col md:flex-row gap-4 p-4 w-full"
                    style={{
                      width: "fit-content",
                      backgroundColor: "rgba(0,0,0,0.05)",
                      borderRadius: "2rem",
                    }}
                  >
                    {/* FRONT SIDE */}
                    <div
                      style={{
                        background: appConfig.cardTemplateUrl
                          ? "none"
                          : "linear-gradient(135deg, #14532d 0%, #052e16 100%)",
                        color: "#ffffff",
                        width: "320px",
                        height: "480px",
                        borderRadius: "32px",
                        position: "relative",
                        overflow: "hidden",
                        boxSizing: "border-box",
                        display: "flex",
                        flexDirection: "column",
                        alignItems: "center",
                        padding: "40px 24px",
                        fontFamily: "Inter, sans-serif",
                      }}
                    >
                      {appConfig.cardTemplateUrl && (
                        <img
                          src={appConfig.cardTemplateUrl}
                          style={{
                            position: "absolute",
                            inset: 0,
                            width: "100%",
                            height: "100%",
                            objectFit: "cover",
                            zIndex: 0,
                          }}
                          crossOrigin="anonymous"
                          alt="Background"
                        />
                      )}
                      {!appConfig.cardTemplateUrl && (
                        <div
                          style={{
                            width: "80px",
                            height: "80px",
                            backgroundColor: "white",
                            borderRadius: "24px",
                            display: "flex",
                            alignItems: "center",
                            justifyContent: "center",
                            padding: "10px",
                            marginBottom: "20px",
                            boxShadow: "0 8px 24px rgba(0,0,0,0.2)",
                            zIndex: 10,
                          }}
                        >
                          <img
                            src={appLogo}
                            style={{
                              width: "100%",
                              height: "100%",
                              objectFit: "contain",
                            }}
                            crossOrigin="anonymous"
                          />
                        </div>
                      )}

                      <div
                        style={{
                          textAlign: "center",
                          marginBottom: "20px",
                          zIndex: 10,
                        }}
                      >
                        <p
                          style={{
                            margin: 0,
                            fontSize: "10px",
                            fontWeight: "900",
                            color: "#facc15",
                            letterSpacing: "0.15em",
                            textTransform: "uppercase",
                            marginBottom: "3px",
                          }}
                        >
                          KARTU SISWA
                        </p>
                        <h3
                          style={{
                            margin: 0,
                            fontSize: "18px",
                            fontWeight: "900",
                            letterSpacing: "0.05em",
                            textTransform: "uppercase",
                            lineHeight: "1.2",
                          }}
                        >
                          MTS NEGERI 2 BOMBANA
                        </h3>
                      </div>

                      <div
                        style={{
                          width: "110px",
                          height: "165px",
                          borderRadius: "24px",
                          backgroundColor: "rgba(255,255,255,0.1)",
                          overflow: "hidden",
                          display: "flex",
                          alignItems: "center",
                          justifyContent: "center",
                          marginBottom: "20px",
                          border: "4px solid rgba(255,255,255,0.1)",
                          zIndex: 20,
                        }}
                      >
                        {selectedStudentCard.foto ? (
                          <img
                            src={selectedStudentCard.foto}
                            style={{
                              width: "100%",
                              height: "100%",
                              objectFit: "cover",
                              display: "block",
                            }}
                            crossOrigin="anonymous"
                          />
                        ) : (
                          <User
                            size={64}
                            style={{ color: "rgba(255,255,255,0.1)" }}
                          />
                        )}
                      </div>

                      <div
                        style={{
                          width: "100%",
                          zIndex: 30,
                          textAlign: "center",
                        }}
                      >
                        <div style={{ marginBottom: "12px" }}>
                          <span
                            style={{
                              fontSize: "9px",
                              fontWeight: "900",
                              color: "rgba(255,255,255,0.4)",
                              textTransform: "uppercase",
                              letterSpacing: "0.1em",
                            }}
                          >
                            Nama Lengkap
                          </span>
                          <h4
                            style={{
                              margin: 0,
                              fontWeight: "900",
                              fontSize: "16px",
                              lineHeight: "1.2",
                              textShadow: "0 2px 4px rgba(0,0,0,0.3)",
                              wordBreak: "break-word",
                            }}
                          >
                            {selectedStudentCard.nama}
                          </h4>
                          <p
                            style={{
                              margin: 0,
                              marginTop: "4px",
                              fontSize: "10px",
                              fontWeight: "normal",
                              color: "rgba(255,255,255,0.8)",
                              letterSpacing: "0.01em",
                            }}
                          >
                            {selectedStudentCard.tempat || "-"}
                            {selectedStudentCard.tgl
                              ? `, ${formatIndoDateNoDay(selectedStudentCard.tgl)}`
                              : ""}
                          </p>
                        </div>

                        <div
                          style={{
                            display: "flex",
                            justifyContent: "center",
                            gap: "20px",
                          }}
                        >
                          <div>
                            <span
                              style={{
                                fontSize: "9px",
                                fontWeight: "900",
                                color: "rgba(255,255,255,0.4)",
                                textTransform: "uppercase",
                                letterSpacing: "0.1em",
                              }}
                            >
                              Kelas
                            </span>
                            <p
                              style={{
                                margin: 0,
                                fontSize: "14px",
                                fontWeight: "900",
                              }}
                            >
                              {selectedStudentCard.kelas}
                            </p>
                          </div>
                          <div>
                            <span
                              style={{
                                fontSize: "9px",
                                fontWeight: "900",
                                color: "rgba(255,255,255,0.4)",
                                textTransform: "uppercase",
                                letterSpacing: "0.1em",
                              }}
                            >
                              NISN
                            </span>
                            <p
                              style={{
                                margin: 0,
                                fontSize: "14px",
                                fontWeight: "900",
                              }}
                            >
                              {selectedStudentCard.nisn}
                            </p>
                          </div>
                        </div>
                      </div>
                    </div>

                    {/* BACK SIDE */}
                    <div
                      style={{
                        background: appConfig.cardTemplateUrl
                          ? "none"
                          : "linear-gradient(135deg, #052e16 0%, #14532d 100%)",
                        color: "#ffffff",
                        width: "320px",
                        height: "480px",
                        borderRadius: "32px",
                        position: "relative",
                        overflow: "hidden",
                        display: "flex",
                        flexDirection: "column",
                        alignItems: "center",
                        padding: "40px 24px",
                        fontFamily: "Inter, sans-serif",
                        justifyContent: "center",
                      }}
                    >
                      {appConfig.cardTemplateUrl && (
                        <img
                          src={appConfig.cardTemplateUrl}
                          style={{
                            position: "absolute",
                            inset: 0,
                            width: "100%",
                            height: "100%",
                            objectFit: "cover",
                            zIndex: 0,
                          }}
                          crossOrigin="anonymous"
                          alt="Background"
                        />
                      )}
                      <div
                        style={{
                          marginBottom: "30px",
                          textAlign: "center",
                          zIndex: 10,
                        }}
                      >
                        <h4
                          style={{
                            fontSize: "14px",
                            fontWeight: "900",
                            textTransform: "uppercase",
                            letterSpacing: "0.1em",
                            marginBottom: "8px",
                          }}
                        >
                          Kartu Presensi Digital
                        </h4>
                        <div
                          style={{
                            height: "2px",
                            width: "40px",
                            background: "#16a34a",
                            margin: "0 auto",
                          }}
                        ></div>
                      </div>

                      <div
                        style={{
                          backgroundColor: "#ffffff",
                          padding: "15px",
                          borderRadius: "24px",
                          boxShadow: "0 8px 24px rgba(0,0,0,0.3)",
                          display: "flex",
                          alignItems: "center",
                          justifyContent: "center",
                        }}
                      >
                        <QRCodeSVG
                          value={selectedStudentCard.nisn}
                          size={150}
                          level="H"
                          includeMargin={false}
                        />
                      </div>
                      <p
                        style={{
                          margin: "15px 0 0 0",
                          fontSize: "14px",
                          fontWeight: "900",
                          textTransform: "uppercase",
                          letterSpacing: "0.05em",
                          color: "#ffffff",
                          textShadow: "0 2px 4px rgba(0,0,0,0.3)",
                          zIndex: 10,
                          textAlign: "center",
                        }}
                      >
                        {selectedStudentCard.nama}
                      </p>
                      <div
                        style={{
                          marginTop: "20px",
                          textAlign: "center",
                          padding: "0 20px",
                        }}
                      >
                        <p
                          style={{
                            fontSize: "10px",
                            fontWeight: "500",
                            opacity: 0.8,
                            lineHeight: "1.6",
                          }}
                        >
                          Simpan kartu ini dengan baik. Gunakan barcode di atas
                          untuk melakukan presensi pada mesin yang tersedia di
                          sekolah.
                        </p>
                        <p
                          style={{
                            fontSize: "9px",
                            fontWeight: "900",
                            marginTop: "20px",
                            opacity: 0.4,
                            letterSpacing: "0.2em",
                          }}
                        >
                          MTS NEGERI 2 BOMBANA
                        </p>
                      </div>
                    </div>
                  </div>

                  <div className="mt-8 flex gap-3 w-full max-w-md">
                    <button
                      onClick={() =>
                        downloadAsPDF(
                          "print-student-card",
                          `Kartu-Siswa-${selectedStudentCard.nisn}.pdf`,
                        )
                      }
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

          {activePanel === "siswa-personal" &&
            (() => {
              const selectedDateAttendance = attendance.find(
                (a) =>
                  a.nisn === session?.uid && a.tanggal === siswaDashboardDate,
              );
              const dayMapLong = [
                "Minggu",
                "Senin",
                "Selasa",
                "Rabu",
                "Kamis",
                "Jumat",
                "Sabtu",
              ];
              const dateObj = new Date(siswaDashboardDate);
              const dayNameLong = dayMapLong[dateObj.getDay()];
              const dateSetting = settings.find((s) => s.hari === dayNameLong);

              return (
                <motion.div
                  key="siswa-personal"
                  initial={{ opacity: 0 }}
                  animate={{ opacity: 1 }}
                  className="space-y-8 pb-32"
                >
                  <div
                    className="flex flex-wrap gap-4 items-center justify-between p-4 rounded-3xl border"
                    style={{
                      backgroundColor: "rgba(255, 255, 255, 0.5)",
                      backdropFilter: "blur(4px)",
                      borderColor: "rgba(255, 255, 255, 0.2)",
                    }}
                  >
                    <div className="flex items-center gap-3">
                      <div className="bg-green-100 p-3 rounded-2xl text-green-700">
                        <Calendar size={24} />
                      </div>
                      <div>
                        <h2 className="text-xl font-black text-green-950 uppercase tracking-tighter">
                          Ringkasan Harian
                        </h2>
                        <p
                          className="text-[10px] font-bold uppercase tracking-widest"
                          style={{ color: "rgba(22, 163, 74, 0.6)" }}
                        >
                          {formatIndoDate(siswaDashboardDate)}
                        </p>
                      </div>
                    </div>
                    <div className="flex items-center gap-2 bg-white p-2 rounded-2xl shadow-sm border">
                      <span className="text-[10px] font-black text-gray-400 uppercase tracking-widest px-2">
                        Pilih Tanggal:
                      </span>
                      <input
                        type="date"
                        value={siswaDashboardDate}
                        onChange={(e) => setSiswaDashboardDate(e.target.value)}
                        className="border-0 focus:ring-0 text-sm font-black text-green-900 bg-transparent"
                      />
                    </div>
                  </div>

                  <div className="grid grid-cols-1 xl:grid-cols-3 gap-8">
                    <div className="xl:col-span-1 space-y-6">
                      <div className="bg-white rounded-[2.5rem] p-8 border border-gray-100 shadow-xl relative overflow-hidden group">
                        <div className="absolute top-0 right-0 w-32 h-32 bg-green-50 rounded-bl-[4rem] -z-10 transition-transform group-hover:scale-110" />
                        <h3 className="text-sm font-black text-green-900 uppercase tracking-[0.2em] mb-8">
                          Status Kehadiran
                        </h3>

                        <div className="space-y-8">
                          <div className="flex items-center justify-between">
                            <div className="flex items-center gap-4">
                              <div className="w-12 h-12 rounded-2xl bg-zinc-50 border border-zinc-100 flex items-center justify-center text-zinc-400">
                                <Clock size={20} />
                              </div>
                              <div>
                                <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest">
                                  Jam Masuk
                                </p>
                                <p className="text-lg font-black text-zinc-900">
                                  {dateSetting?.masuk || "--:--"}
                                </p>
                              </div>
                            </div>
                            {selectedDateAttendance?.jam && (
                              <div className="text-right">
                                <p className="text-[10px] font-black text-green-600 uppercase tracking-widest">
                                  Absen Masuk
                                </p>
                                <p className="text-lg font-black text-green-700">
                                  {selectedDateAttendance.jam}
                                </p>
                              </div>
                            )}
                          </div>

                          <div className="flex items-center justify-between">
                            <div className="flex items-center gap-4">
                              <div className="w-12 h-12 rounded-2xl bg-zinc-50 border border-zinc-100 flex items-center justify-center text-zinc-400">
                                <LogOut size={20} className="rotate-180" />
                              </div>
                              <div>
                                <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest">
                                  Jam Pulang
                                </p>
                                <p className="text-lg font-black text-zinc-900">
                                  {dateSetting?.pulang || "--:--"}
                                </p>
                              </div>
                            </div>
                            {selectedDateAttendance?.jamPulang && (
                              <div className="text-right">
                                <p className="text-[10px] font-black text-blue-600 uppercase tracking-widest">
                                  Absen Pulang
                                </p>
                                <p className="text-lg font-black text-blue-700">
                                  {selectedDateAttendance.jamPulang}
                                </p>
                              </div>
                            )}
                          </div>

                          <div className="pt-6 border-t border-gray-50">
                            <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest mb-3">
                              Status Akhir
                            </p>
                            <div
                              className={`py-4 rounded-3xl text-center border-2 transition-all`}
                              style={{
                                ...(selectedDateAttendance?.status === "Hadir"
                                  ? {
                                      backgroundColor: "#f0fdf4",
                                      borderColor: "#bbf7d0",
                                      color: "#166534",
                                    }
                                  : selectedDateAttendance?.status === "Izin"
                                    ? {
                                        backgroundColor: "#eff6ff",
                                        borderColor: "#bfdbfe",
                                        color: "#1e40af",
                                      }
                                    : selectedDateAttendance?.status === "Sakit"
                                      ? {
                                          backgroundColor: "#fffbeb",
                                          borderColor: "#fde68a",
                                          color: "#92400e",
                                        }
                                      : selectedDateAttendance?.status ===
                                          "Alfa"
                                        ? {
                                            backgroundColor: "#fef2f2",
                                            borderColor: "#fecaca",
                                            color: "#991b1b",
                                          }
                                        : {
                                            backgroundColor: "#f9fafb",
                                            borderColor: "#e5e7eb",
                                            color: "#9ca3af",
                                          }),
                              }}
                            >
                              <span className="text-2xl font-black uppercase tracking-widest block">
                                {selectedDateAttendance?.status ||
                                  "Belum Absen"}
                              </span>
                              {selectedDateAttendance?.status === "Hadir" && (
                                <p
                                  className="text-[10px] font-black mt-1 uppercase tracking-widest"
                                  style={{ color: "rgba(22, 163, 74, 0.6)" }}
                                >
                                  Selamat Belajar!
                                </p>
                              )}
                              {selectedDateAttendance?.terlambat &&
                              selectedDateAttendance.terlambat > 0 ? (
                                <p
                                  className="text-[10px] font-black mt-2 px-3 py-1 rounded-full inline-block"
                                  style={{
                                    backgroundColor: "#fee2e2",
                                    color: "#b91c1c",
                                  }}
                                >
                                  Terlambat {selectedDateAttendance.terlambat}{" "}
                                  Menit
                                </p>
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
                              <circle
                                cx="64"
                                cy="64"
                                r="56"
                                stroke="currentColor"
                                strokeWidth="12"
                                fill="transparent"
                                className="text-gray-100"
                              />
                              <circle
                                cx="64"
                                cy="64"
                                r="56"
                                stroke="currentColor"
                                strokeWidth="12"
                                fill="transparent"
                                strokeDasharray={351.8}
                                strokeDashoffset={
                                  351.8 - (351.8 * getSiswaPercentage) / 100
                                }
                                strokeLinecap="round"
                                className="text-green-600 transition-all duration-1000"
                              />
                            </svg>
                            <div className="absolute flex flex-col items-center">
                              <span className="text-2xl font-black text-green-900">
                                {getSiswaPercentage}%
                              </span>
                              <span className="text-[8px] font-bold text-gray-400 uppercase">
                                Hadir
                              </span>
                            </div>
                          </div>
                          <div className="flex-1 w-full space-y-3">
                            <div className="p-3 bg-gray-50 rounded-2xl flex items-center justify-between">
                              <span className="text-xs font-bold text-gray-500">
                                Total Absen
                              </span>
                              <span className="text-sm font-black text-green-900">
                                {
                                  attendance.filter(
                                    (a) =>
                                      a.nisn === session?.uid &&
                                      a.tanggal.startsWith(
                                        new Date().toISOString().slice(0, 7),
                                      ),
                                  ).length
                                }
                              </span>
                            </div>
                            <div className="p-3 bg-green-50 rounded-2xl flex items-center justify-between">
                              <span className="text-xs font-bold text-green-700">
                                Hadir
                              </span>
                              <span className="text-sm font-black text-green-800">
                                {
                                  attendance.filter(
                                    (a) =>
                                      a.nisn === session?.uid &&
                                      a.tanggal.startsWith(
                                        new Date().toISOString().slice(0, 7),
                                      ) &&
                                      a.status === "Hadir",
                                  ).length
                                }
                              </span>
                            </div>
                          </div>
                        </div>
                        <p className="mt-6 text-[10px] text-gray-400 font-medium italic text-center">
                          "Kedisiplinan adalah modal utama meraih cita-cita."
                        </p>
                      </div>

                      {/* PUSAT NOTIFIKASI ORANG TUA */}
                      <div className="bg-white rounded-[2.5rem] p-8 border border-gray-100 shadow-xl relative overflow-hidden">
                        <div className="absolute top-0 right-0 w-32 h-32 bg-amber-50 rounded-bl-[4rem] -z-10" />

                        <div className="flex items-center gap-3 mb-6">
                          <div className="bg-amber-100 p-3 rounded-2xl text-amber-700 animate-pulse">
                            <Bell size={20} />
                          </div>
                          <div>
                            <h3 className="text-sm font-black text-amber-950 uppercase tracking-wider">
                              Notifikasi Orang Tua
                            </h3>
                            <p className="text-[9px] font-bold text-gray-400 uppercase tracking-widest">
                              Real-Time Tracker
                            </p>
                          </div>
                        </div>

                        <p className="text-xs text-gray-500 font-semibold leading-relaxed mb-6">
                          Dapatkan notifikasi instan langsung di perangkat HP
                          atau Laptop Anda ketika anak Anda terdeteksi melakukan
                          presensi <strong>MASUK</strong> atau{" "}
                          <strong>PULANG</strong> di madrasah.
                        </p>

                        {isInIframe && (
                          <div className="p-4 bg-amber-50/75 border border-amber-200/60 text-amber-950 rounded-2xl text-xs font-sans mb-5 leading-normal space-y-1">
                            <p className="font-black text-amber-900 uppercase text-[10px]">
                              ⚠️ Anda Sedang di Dalam Preview
                            </p>
                            <p className="text-[11px] font-medium text-amber-800">
                              Browser melarang permintaan izin Notifikasi dari
                              dalam frame visual ini.
                            </p>
                            <p className="text-[11px] font-bold text-amber-950">
                              Untuk mengaktifkan Notifikasi Sistem HP, klik
                              tombol{" "}
                              <span className="underline">
                                "Open in new tab"
                              </span>{" "}
                              di pojok kanan atas preview Anda terlebih dahulu.
                            </p>
                          </div>
                        )}

                        <div className="space-y-4">
                          {/* Status Izin Notifikasi Browser */}
                          <div className="p-4 bg-zinc-50 rounded-2xl border border-zinc-100 flex items-center justify-between">
                            <div>
                              <p className="text-[10px] font-black text-zinc-400 uppercase tracking-widest leading-none">
                                Izin Browser
                              </p>
                              <p className="text-xs font-bold text-zinc-800 mt-1 capitalize">
                                {notifPermission === "granted"
                                  ? "Diizinkan ✓"
                                  : notifPermission === "denied"
                                    ? "Ditolak ✗"
                                    : "Belum Ditentukan"}
                              </p>
                            </div>
                            {notifPermission !== "granted" ? (
                              <button
                                type="button"
                                onClick={async () => {
                                  if (
                                    typeof window === "undefined" ||
                                    !("Notification" in window)
                                  ) {
                                    alert(
                                      "Browser Anda tidak mendukung Web Notifications.",
                                    );
                                    return;
                                  }
                                  try {
                                    const permission =
                                      await Notification.requestPermission();
                                    setNotifPermission(permission);
                                    if (permission === "granted") {
                                      setParentNotifEnabled(true);
                                      localStorage.setItem(
                                        "parent_notif_enabled",
                                        "true",
                                      );
                                      triggerSuccess(
                                        "NOTIFIKASI AKTIF",
                                        "Selamat! Notifikasi kehadiran berhasil diaktifkan.",
                                      );
                                      try {
                                        new Notification(
                                          "🔔 Notifikasi Orang Tua Aktif",
                                          {
                                            body: "SIGAP: Anda sekarang akan menerima update kehadiran real-time di perangkat ini.",
                                            icon:
                                              appConfig.logoUrl ||
                                              "/favicon.ico",
                                          },
                                        );
                                      } catch (e) {
                                        console.error(e);
                                      }
                                    } else {
                                      alert(
                                        `Izin notifikasi ditolak/tidak diberikan (${permission}). Silakan periksa pengaturan situs di browser Anda.`,
                                      );
                                    }
                                  } catch (err) {
                                    console.warn(
                                      "Notification.requestPermission failed:",
                                      err,
                                    );
                                    // Fallback handling for iframe block
                                    alert(
                                      "Permintaan izin diblokir karena dijalankan dalam preview frame.\n\n" +
                                        "Silakan klik tombol 'Open in new tab' di pojok kanan atas preview untuk membuka SIGAP secara penuh, " +
                                        "lalu klik kembali 'Izinkan' untuk mengaktifkannya secara resmi.\n\n" +
                                        "Namun, kami telah mengaktifkan simulasi notifikasi di dalam aplikasi (In-App Toast Alerts) untuk memfasilitasi pengetesan Anda!",
                                    );
                                    setParentNotifEnabled(true);
                                    localStorage.setItem(
                                      "parent_notif_enabled",
                                      "true",
                                    );
                                    triggerSuccess(
                                      "SIMULASI DIAKTIFKAN",
                                      "Status aktif secara simulasi (In-App Toast notifications).",
                                    );
                                  }
                                }}
                                className="bg-amber-600 hover:bg-amber-700 text-white font-black uppercase text-[9px] tracking-wider px-3.5 py-2 rounded-xl transition-all shadow-sm active:scale-95 cursor-pointer"
                              >
                                Izinkan
                              </button>
                            ) : (
                              <span className="bg-green-100 text-green-700 border border-green-200 font-bold uppercase text-[9px] px-2.5 py-1 rounded-lg font-sans">
                                Aktif
                              </span>
                            )}
                          </div>

                          {/* Toggle Notifikasi Kehadiran */}
                          <div className="p-4 bg-zinc-50 rounded-2xl border border-zinc-100 flex items-center justify-between">
                            <div>
                              <p className="text-[10px] font-black text-zinc-400 uppercase tracking-widest leading-none">
                                Status Notifikasi
                              </p>
                              <p className="text-xs font-bold text-zinc-800 mt-1">
                                {parentNotifEnabled
                                  ? "Aktif & Siaga 🔔"
                                  : "Nonaktif 🔕"}
                              </p>
                            </div>
                            <button
                              type="button"
                              onClick={async () => {
                                const nextVal = !parentNotifEnabled;
                                setParentNotifEnabled(nextVal);
                                localStorage.setItem(
                                  "parent_notif_enabled",
                                  nextVal ? "true" : "false",
                                );

                                if (nextVal) {
                                  if (
                                    typeof window !== "undefined" &&
                                    "Notification" in window
                                  ) {
                                    try {
                                      const p =
                                        await Notification.requestPermission();
                                      setNotifPermission(p);
                                      if (p !== "granted") {
                                        triggerSuccess(
                                          "MODE IN-APP AKTIF",
                                          "Status aktif. Notifikasi akan dimunculkan di dalam aplikasi karena izin sistem diblokir/ditolak.",
                                        );
                                      } else {
                                        triggerSuccess(
                                          "NOTIFIKASI AKTIF",
                                          "Anda akan menerima notifikasi otomatis ketika anak Anda melakukan presensi.",
                                        );
                                      }
                                    } catch (err) {
                                      console.warn(
                                        "Request failed in toggle:",
                                        err,
                                      );
                                      triggerSuccess(
                                        "MODE IN-APP AKTIF",
                                        "Status sistem aktif. Notifikasi akan muncul sebagai notifikasi langsung di layar aplikasi.",
                                      );
                                    }
                                  } else {
                                    triggerSuccess(
                                      "MODE IN-APP AKTIF",
                                      "Status sistem aktif. Perangkat Anda tidak mendukung notifikasi sistem, menggunakan mode dalam aplikasi.",
                                    );
                                  }
                                } else {
                                  triggerSuccess(
                                    "NOTIFIKASI DIMATIKAN",
                                    "Notifikasi kehadiran telah dinonaktifkan.",
                                  );
                                }
                              }}
                              className={`w-12 h-6 rounded-full transition-colors relative cursor-pointer ${
                                parentNotifEnabled
                                  ? "bg-amber-600"
                                  : "bg-zinc-300"
                              }`}
                            >
                              <span
                                className={`absolute top-1 left-1 bg-white w-4 h-4 rounded-full transition-transform ${
                                  parentNotifEnabled
                                    ? "translate-x-6"
                                    : "translate-x-0"
                                }`}
                              />
                            </button>
                          </div>

                          {/* Test Notification and Guidance Row */}
                          <div className="pt-2 flex gap-2">
                            <button
                              type="button"
                              onClick={() => {
                                if (typeof window === "undefined") return;

                                const hasNativePermission =
                                  "Notification" in window &&
                                  Notification.permission === "granted";
                                if (!hasNativePermission) {
                                  // Fallback to test with in-app notification directly
                                  triggerSuccess(
                                    "📍 TEST PRESENSI (IN-APP)",
                                    "SIGAP: Test update berhasil dikirim di layar ini.",
                                  );
                                } else {
                                  triggerParentNotification(
                                    "🔔 SIGAP: Tes Koneksi",
                                    `Ini adalah contoh notifikasi kehadiran dari aplikasi SIGAP Madrasah.`,
                                  );
                                }
                              }}
                              className="flex-1 bg-zinc-100 hover:bg-zinc-200 text-zinc-800 border border-zinc-200/60 font-black uppercase text-[9px] tracking-wider py-2.5 rounded-xl transition-all shadow-sm cursor-pointer text-center"
                            >
                              Uji Notifikasi
                            </button>
                          </div>

                          <div className="p-3 bg-amber-50/50 rounded-2xl border border-amber-100/60 text-amber-950 font-sans mt-3">
                            <p className="text-[10px] font-black uppercase text-amber-900 mb-1">
                              💡 Tips Latar Belakang (Penting)
                            </p>
                            <p className="text-[9px] text-amber-800 leading-normal font-medium">
                              Agar notifikasi tetap masuk saat HP terkunci atau
                              sedang membuka aplikasi lain:
                              <br />
                              1. Biarkan tab SIGAP tetap terbuka di browser
                              latar belakang Anda.
                              <br />
                              2. Pastikan izin notifikasi browser HP Anda dalam
                              posisi "Izinkan" dan set hemat baterai ke "Tanpa
                              Batasan" jika didukung.
                            </p>
                          </div>
                        </div>
                      </div>
                    </div>

                    <div className="xl:col-span-2 space-y-8">
                      <div
                        className="flex flex-col items-center w-full p-8 rounded-[3rem] border shadow-sm relative group"
                        style={{
                          backgroundColor: "rgba(255, 255, 255, 0.4)",
                          backdropFilter: "blur(12px)",
                          borderColor: "rgba(255, 255, 255, 0.5)",
                        }}
                      >
                        <div className="absolute top-6 right-6 opacity-20 group-hover:opacity-100 transition-opacity">
                          <CreditCard className="text-green-900" size={32} />
                        </div>
                        <div
                          id="personal-student-card"
                          className="flex flex-col lg:flex-row gap-4 p-4 w-fit max-w-full overflow-auto scale-75 md:scale-95 lg:scale-100 origin-top"
                          style={{
                            backgroundColor: "rgba(0,0,0,0.05)",
                            borderRadius: "2.5rem",
                            boxShadow: "0 25px 50px -12px rgba(0, 0, 0, 0.25)",
                          }}
                        >
                          {/* FRONT SIDE */}
                          <div
                            style={{
                              background: appConfig.cardTemplateUrl
                                ? "none"
                                : "linear-gradient(135deg, #14532d 0%, #052e16 100%)",
                              color: "#ffffff",
                              width: "320px",
                              height: "480px",
                              borderRadius: "32px",
                              position: "relative",
                              overflow: "hidden",
                              boxSizing: "border-box",
                              display: "flex",
                              flexDirection: "column",
                              alignItems: "center",
                              padding: "40px 24px",
                              fontFamily: "Inter, sans-serif",
                            }}
                          >
                            {appConfig.cardTemplateUrl && (
                              <img
                                src={appConfig.cardTemplateUrl}
                                style={{
                                  position: "absolute",
                                  inset: 0,
                                  width: "100%",
                                  height: "100%",
                                  objectFit: "cover",
                                  zIndex: 0,
                                }}
                                crossOrigin="anonymous"
                                alt="Background"
                              />
                            )}
                            {!appConfig.cardTemplateUrl && (
                              <div
                                style={{
                                  width: "80px",
                                  height: "80px",
                                  backgroundColor: "white",
                                  borderRadius: "24px",
                                  display: "flex",
                                  alignItems: "center",
                                  justifyContent: "center",
                                  padding: "10px",
                                  marginBottom: "20px",
                                  boxShadow: "0 8px 24px rgba(0,0,0,0.2)",
                                  zIndex: 10,
                                }}
                              >
                                <img
                                  src={appLogo}
                                  style={{
                                    width: "100%",
                                    height: "100%",
                                    objectFit: "contain",
                                  }}
                                  crossOrigin="anonymous"
                                />
                              </div>
                            )}

                            <div
                              style={{
                                textAlign: "center",
                                marginBottom: "20px",
                                zIndex: 10,
                              }}
                            >
                              <p
                                style={{
                                  margin: 0,
                                  fontSize: "10px",
                                  fontWeight: "900",
                                  color: "#facc15",
                                  letterSpacing: "0.15em",
                                  textTransform: "uppercase",
                                  marginBottom: "3px",
                                }}
                              >
                                KARTU SISWA
                              </p>
                              <h3
                                style={{
                                  margin: 0,
                                  fontSize: "18px",
                                  fontWeight: "900",
                                  letterSpacing: "0.05em",
                                  textTransform: "uppercase",
                                  lineHeight: "1.2",
                                }}
                              >
                                MTS NEGERI 2 BOMBANA
                              </h3>
                            </div>

                            <div
                              style={{
                                width: "110px",
                                height: "165px",
                                borderRadius: "24px",
                                backgroundColor: "rgba(255,255,255,0.1)",
                                overflow: "hidden",
                                display: "flex",
                                alignItems: "center",
                                justifyContent: "center",
                                marginBottom: "20px",
                                border: "4px solid rgba(255,255,255,0.1)",
                                zIndex: 20,
                              }}
                            >
                              {currentSiswaData?.foto ? (
                                <img
                                  src={currentSiswaData.foto}
                                  style={{
                                    width: "100%",
                                    height: "100%",
                                    objectFit: "cover",
                                    display: "block",
                                  }}
                                  crossOrigin="anonymous"
                                />
                              ) : (
                                <User
                                  size={64}
                                  style={{ color: "rgba(255,255,255,0.1)" }}
                                />
                              )}
                            </div>

                            <div
                              style={{
                                width: "100%",
                                zIndex: 30,
                                textAlign: "center",
                              }}
                            >
                              <div style={{ marginBottom: "12px" }}>
                                <span
                                  style={{
                                    fontSize: "9px",
                                    fontWeight: "900",
                                    color: "rgba(255,255,255,0.4)",
                                    textTransform: "uppercase",
                                    letterSpacing: "0.1em",
                                  }}
                                >
                                  Nama Lengkap
                                </span>
                                <h4
                                  style={{
                                    margin: 0,
                                    fontWeight: "900",
                                    fontSize: "16px",
                                    lineHeight: "1.2",
                                    textShadow: "0 2px 4px rgba(0,0,0,0.3)",
                                    wordBreak: "break-word",
                                  }}
                                >
                                  {session?.name || "Siswa"}
                                </h4>
                                <p
                                  style={{
                                    margin: 0,
                                    marginTop: "4px",
                                    fontSize: "10px",
                                    fontWeight: "normal",
                                    color: "rgba(255,255,255,0.8)",
                                    letterSpacing: "0.01em",
                                  }}
                                >
                                  {currentSiswaData?.tempat || "-"}
                                  {currentSiswaData?.tgl
                                    ? `, ${formatIndoDateNoDay(currentSiswaData?.tgl)}`
                                    : ""}
                                </p>
                              </div>

                              <div
                                style={{
                                  display: "flex",
                                  justifyContent: "center",
                                  gap: "20px",
                                }}
                              >
                                <div>
                                  <span
                                    style={{
                                      fontSize: "9px",
                                      fontWeight: "900",
                                      color: "rgba(255,255,255,0.4)",
                                      textTransform: "uppercase",
                                      letterSpacing: "0.1em",
                                    }}
                                  >
                                    Kelas
                                  </span>
                                  <p
                                    style={{
                                      margin: 0,
                                      fontSize: "14px",
                                      fontWeight: "900",
                                    }}
                                  >
                                    {session?.kelas}
                                  </p>
                                </div>
                                <div>
                                  <span
                                    style={{
                                      fontSize: "9px",
                                      fontWeight: "900",
                                      color: "rgba(255,255,255,0.4)",
                                      textTransform: "uppercase",
                                      letterSpacing: "0.1em",
                                    }}
                                  >
                                    NISN
                                  </span>
                                  <p
                                    style={{
                                      margin: 0,
                                      fontSize: "14px",
                                      fontWeight: "900",
                                    }}
                                  >
                                    {session?.uid}
                                  </p>
                                </div>
                              </div>
                            </div>
                          </div>

                          {/* BACK SIDE */}
                          <div
                            style={{
                              background: appConfig.cardTemplateUrl
                                ? "none"
                                : "linear-gradient(135deg, #052e16 0%, #14532d 100%)",
                              color: "#ffffff",
                              width: "320px",
                              height: "480px",
                              borderRadius: "32px",
                              position: "relative",
                              overflow: "hidden",
                              display: "flex",
                              flexDirection: "column",
                              alignItems: "center",
                              padding: "40px 24px",
                              fontFamily: "Inter, sans-serif",
                              justifyContent: "center",
                            }}
                          >
                            {appConfig.cardTemplateUrl && (
                              <img
                                src={appConfig.cardTemplateUrl}
                                style={{
                                  position: "absolute",
                                  inset: 0,
                                  width: "100%",
                                  height: "100%",
                                  objectFit: "cover",
                                  zIndex: 0,
                                }}
                                crossOrigin="anonymous"
                                alt="Background"
                              />
                            )}
                            <div
                              style={{
                                marginBottom: "30px",
                                textAlign: "center",
                                zIndex: 10,
                              }}
                            >
                              <h4
                                style={{
                                  fontSize: "14px",
                                  fontWeight: "900",
                                  textTransform: "uppercase",
                                  letterSpacing: "0.1em",
                                  marginBottom: "8px",
                                }}
                              >
                                Kartu Presensi Digital
                              </h4>
                              <div
                                style={{
                                  height: "2px",
                                  width: "40px",
                                  background: "#16a34a",
                                  margin: "0 auto",
                                }}
                              ></div>
                            </div>

                            <div
                              style={{
                                backgroundColor: "#ffffff",
                                padding: "15px",
                                borderRadius: "24px",
                                boxShadow: "0 8px 24px rgba(0,0,0,0.3)",
                                display: "flex",
                                alignItems: "center",
                                justifyContent: "center",
                              }}
                            >
                              <QRCodeSVG
                                value={session?.uid || ""}
                                size={150}
                                level="H"
                                includeMargin={false}
                              />
                            </div>
                            <p
                              style={{
                                margin: "15px 0 0 0",
                                fontSize: "14px",
                                fontWeight: "900",
                                textTransform: "uppercase",
                                letterSpacing: "0.05em",
                                color: "#ffffff",
                                textShadow: "0 2px 4px rgba(0,0,0,0.3)",
                                zIndex: 10,
                                textAlign: "center",
                              }}
                            >
                              {session?.name || "Siswa"}
                            </p>
                            <div
                              style={{
                                marginTop: "20px",
                                textAlign: "center",
                                padding: "0 20px",
                              }}
                            >
                              <p
                                style={{
                                  fontSize: "10px",
                                  fontWeight: "500",
                                  opacity: 0.8,
                                  lineHeight: "1.6",
                                }}
                              >
                                Simpan kartu ini dengan baik. Gunakan barcode di
                                atas untuk melakukan presensi pada mesin yang
                                tersedia di sekolah.
                              </p>
                              <p
                                style={{
                                  fontSize: "9px",
                                  fontWeight: "900",
                                  marginTop: "20px",
                                  opacity: 0.4,
                                  letterSpacing: "0.2em",
                                }}
                              >
                                MTS NEGERI 2 BOMBANA
                              </p>
                            </div>
                          </div>
                        </div>

                        <div className="mt-8 flex gap-4 w-full justify-center">
                          <button
                            onClick={() =>
                              downloadAsPDF(
                                "personal-student-card",
                                `Kartu-${session?.uid}.pdf`,
                              )
                            }
                            className="bg-green-800 text-white px-8 py-4 rounded-2xl font-black text-xs uppercase tracking-widest shadow-xl flex items-center justify-center gap-2 hover:bg-green-700 active:scale-95 transition-all"
                          >
                            <Download size={18} /> Simpan Kartu (PDF)
                          </button>
                        </div>
                      </div>

                      <div className="bg-white rounded-[2rem] shadow-sm border border-gray-100 overflow-hidden">
                        <div className="p-6 bg-gray-50 border-b border-gray-100 flex items-center justify-between">
                          <h3 className="text-sm font-black text-green-900 uppercase tracking-widest">
                            Riwayat Kehadiran Terakhir
                          </h3>
                          <div className="bg-white px-3 py-1 rounded-full border border-gray-200 text-[9px] font-black uppercase text-gray-400">
                            7 Data Terakhir
                          </div>
                        </div>
                        <div className="overflow-x-auto">
                          <table className="w-full text-xs text-left">
                            <thead>
                              <tr
                                style={{
                                  backgroundColor: "rgba(249, 250, 251, 0.5)",
                                }}
                              >
                                <th className="px-6 py-4 font-black">
                                  TANGGAL
                                </th>
                                <th className="px-6 py-4 font-black">
                                  JAM MASUK
                                </th>
                                <th className="px-6 py-4 font-black">
                                  JAM PULANG
                                </th>
                                <th className="px-6 py-4 font-black">STATUS</th>
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-gray-50">
                              {attendance
                                .filter((a) => a.nisn === session?.uid)
                                .sort((a, b) =>
                                  b.tanggal.localeCompare(a.tanggal),
                                )
                                .slice(0, 7)
                                .map((a, i) => (
                                  <tr
                                    key={i}
                                    className="hover:bg-gray-50 transition-colors"
                                  >
                                    <td className="px-6 py-4 font-bold text-gray-500 text-nowrap">
                                      {formatIndoDate(a.tanggal)}
                                    </td>
                                    <td className="px-6 py-4 font-bold text-green-600">
                                      {a.jam || "-"}
                                    </td>
                                    <td className="px-6 py-4 font-bold text-blue-600">
                                      {a.jamPulang || "-"}
                                    </td>
                                    <td className="px-6 py-4">
                                      <span
                                        className={`px-2 py-0.5 rounded-full text-[10px] font-black ${
                                          a.status === "Hadir"
                                            ? "bg-green-100 text-green-700"
                                            : a.status === "Alfa"
                                              ? "bg-red-100 text-red-700"
                                              : "bg-amber-100 text-amber-700"
                                        }`}
                                      >
                                        {a.status}
                                      </span>
                                    </td>
                                  </tr>
                                ))}
                              {attendance.filter((a) => a.nisn === session?.uid)
                                .length === 0 && (
                                <tr>
                                  <td
                                    colSpan={4}
                                    className="px-6 py-12 text-center text-gray-400 font-bold uppercase tracking-widest text-[10px]"
                                  >
                                    Belum ada riwayat kehadiran.
                                  </td>
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
        {renderWarningToast()}
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
                <h2 className="text-2xl md:text-4xl font-black text-green-950 uppercase tracking-tighter">
                  DATA BERHASIL DISIMPAN
                </h2>
                <p className="text-xs md:text-sm font-bold text-zinc-500 mt-2 uppercase tracking-widest">
                  Sinkronisasi Real-time Berhasil
                </p>
              </motion.div>
            </motion.div>
          )}
        </AnimatePresence>
        <ConfirmModal />
        <ResetAbsensiModal />
        <StudentAttendanceDetailModal />
        <MonthlyMapelAttendanceModal />
        <MassDownloadConfirmModal />

        {/* Off-screen Mass Cards Renderer */}
        {massDownloadSiswa.length > 0 && (
          <div
            style={{
              position: "fixed",
              top: "-9999px",
              left: "-9999px",
              width: "auto",
              height: "auto",
              pointerEvents: "none",
              zIndex: -9999,
              backgroundColor: "white",
              display: "flex",
              flexDirection: "column",
              gap: "24px",
            }}
          >
            {massDownloadSiswa.map((student) => {
              return (
                <div
                  key={student.nisn}
                  id={`mass-print-student-card-${student.nisn}`}
                  className="flex flex-col md:flex-row gap-4 p-4"
                  style={{
                    width: "fit-content",
                    backgroundColor: "#f4f4f5",
                    borderRadius: "32px",
                    display: "flex",
                  }}
                >
                  {/* FRONT SIDE */}
                  <div
                    style={{
                      background: appConfig.cardTemplateUrl
                        ? "none"
                        : "linear-gradient(135deg, #14532d 0%, #052e16 100%)",
                      color: "#ffffff",
                      width: "320px",
                      height: "480px",
                      borderRadius: "32px",
                      position: "relative",
                      overflow: "hidden",
                      display: "flex",
                      flexDirection: "column",
                      alignItems: "center",
                      padding: "40px 24px",
                      fontFamily: "Inter, sans-serif",
                      boxSizing: "border-box",
                    }}
                  >
                    {appConfig.cardTemplateUrl && (
                      <img
                        src={appConfig.cardTemplateUrl}
                        style={{
                          position: "absolute",
                          inset: 0,
                          width: "100%",
                          height: "100%",
                          objectFit: "cover",
                          zIndex: 0,
                        }}
                        crossOrigin="anonymous"
                        alt="Background"
                      />
                    )}
                    {!appConfig.cardTemplateUrl && (
                      <div
                        style={{
                          width: "80px",
                          height: "80px",
                          backgroundColor: "white",
                          borderRadius: "24px",
                          display: "flex",
                          alignItems: "center",
                          justifyContent: "center",
                          padding: "10px",
                          marginBottom: "20px",
                          boxShadow: "0 8px 24px rgba(0,0,0,0.2)",
                          zIndex: 10,
                        }}
                      >
                        <img
                          src={appLogo}
                          style={{
                            width: "100%",
                            height: "100%",
                            objectFit: "contain",
                          }}
                          crossOrigin="anonymous"
                        />
                      </div>
                    )}

                    <div
                      style={{
                        textAlign: "center",
                        marginBottom: "20px",
                        zIndex: 10,
                      }}
                    >
                      <p
                        style={{
                          margin: 0,
                          fontSize: "10px",
                          fontWeight: "900",
                          color: "#facc15",
                          letterSpacing: "0.15em",
                          textTransform: "uppercase",
                          marginBottom: "3px",
                        }}
                      >
                        KARTU SISWA
                      </p>
                      <h3
                        style={{
                          margin: 0,
                          fontSize: "18px",
                          fontWeight: "900",
                          letterSpacing: "0.05em",
                          textTransform: "uppercase",
                          lineHeight: "1.2",
                        }}
                      >
                        MTS NEGERI 2 BOMBANA
                      </h3>
                    </div>

                    <div
                      style={{
                        width: "110px",
                        height: "165px",
                        borderRadius: "24px",
                        backgroundColor: "rgba(255,255,255,0.1)",
                        overflow: "hidden",
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "center",
                        marginBottom: "20px",
                        border: "4px solid rgba(255,255,255,0.1)",
                        zIndex: 20,
                      }}
                    >
                      {student.foto ? (
                        <img
                          src={student.foto}
                          style={{
                            width: "100%",
                            height: "100%",
                            objectFit: "cover",
                            display: "block",
                          }}
                          crossOrigin="anonymous"
                        />
                      ) : (
                        <User
                          size={64}
                          style={{ color: "rgba(255,255,255,0.1)" }}
                        />
                      )}
                    </div>

                    <div
                      style={{ width: "100%", zIndex: 30, textAlign: "center" }}
                    >
                      <div style={{ marginBottom: "12px" }}>
                        <span
                          style={{
                            fontSize: "9px",
                            fontWeight: "900",
                            color: "rgba(255,255,255,0.4)",
                            textTransform: "uppercase",
                            letterSpacing: "0.1em",
                          }}
                        >
                          Nama Lengkap
                        </span>
                        <h4
                          style={{
                            margin: 0,
                            fontWeight: "900",
                            fontSize: "16px",
                            lineHeight: "1.2",
                            textShadow: "0 2px 4px rgba(0,0,0,0.3)",
                            wordBreak: "break-word",
                          }}
                        >
                          {student.nama}
                        </h4>
                        <p
                          style={{
                            margin: 0,
                            marginTop: "4px",
                            fontSize: "10px",
                            fontWeight: "normal",
                            color: "rgba(255,255,255,0.8)",
                            letterSpacing: "0.01em",
                          }}
                        >
                          {student.tempat || "-"}
                          {student.tgl
                            ? `, ${formatIndoDateNoDay(student.tgl)}`
                            : ""}
                        </p>
                      </div>

                      <div
                        style={{
                          display: "flex",
                          justifyContent: "center",
                          gap: "20px",
                        }}
                      >
                        <div>
                          <span
                            style={{
                              fontSize: "9px",
                              fontWeight: "900",
                              color: "rgba(255,255,255,0.4)",
                              textTransform: "uppercase",
                              letterSpacing: "0.1em",
                            }}
                          >
                            Kelas
                          </span>
                          <p
                            style={{
                              margin: 0,
                              fontSize: "14px",
                              fontWeight: "900",
                            }}
                          >
                            {student.kelas}
                          </p>
                        </div>
                        <div>
                          <span
                            style={{
                              fontSize: "9px",
                              fontWeight: "900",
                              color: "rgba(255,255,255,0.4)",
                              textTransform: "uppercase",
                              letterSpacing: "0.1em",
                            }}
                          >
                            NISN
                          </span>
                          <p
                            style={{
                              margin: 0,
                              fontSize: "14px",
                              fontWeight: "900",
                            }}
                          >
                            {student.nisn}
                          </p>
                        </div>
                      </div>
                    </div>
                  </div>

                  {/* BACK SIDE */}
                  <div
                    style={{
                      background: appConfig.cardTemplateUrl
                        ? "none"
                        : "linear-gradient(135deg, #052e16 0%, #14532d 100%)",
                      color: "#ffffff",
                      width: "320px",
                      height: "480px",
                      borderRadius: "32px",
                      position: "relative",
                      overflow: "hidden",
                      display: "flex",
                      flexDirection: "column",
                      alignItems: "center",
                      padding: "40px 24px",
                      fontFamily: "Inter, sans-serif",
                      justifyContent: "center",
                      boxSizing: "border-box",
                    }}
                  >
                    {appConfig.cardTemplateUrl && (
                      <img
                        src={appConfig.cardTemplateUrl}
                        style={{
                          position: "absolute",
                          inset: 0,
                          width: "100%",
                          height: "100%",
                          objectFit: "cover",
                          zIndex: 0,
                        }}
                        crossOrigin="anonymous"
                        alt="Background"
                      />
                    )}
                    <div
                      style={{
                        marginBottom: "30px",
                        textAlign: "center",
                        zIndex: 10,
                      }}
                    >
                      <h4
                        style={{
                          fontSize: "14px",
                          fontWeight: "900",
                          textTransform: "uppercase",
                          letterSpacing: "0.1em",
                          marginBottom: "8px",
                        }}
                      >
                        Kartu Presensi Digital
                      </h4>
                      <div
                        style={{
                          height: "2px",
                          width: "40px",
                          background: "#16a34a",
                          margin: "0 auto",
                        }}
                      ></div>
                    </div>

                    <div
                      style={{
                        backgroundColor: "#ffffff",
                        padding: "15px",
                        borderRadius: "24px",
                        boxShadow: "0 8px 24px rgba(0,0,0,0.3)",
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "center",
                      }}
                    >
                      <QRCodeSVG
                        value={student.nisn}
                        size={150}
                        level="H"
                        includeMargin={false}
                      />
                    </div>
                    <p
                      style={{
                        margin: "15px 0 0 0",
                        fontSize: "14px",
                        fontWeight: "900",
                        textTransform: "uppercase",
                        letterSpacing: "0.05em",
                        color: "#ffffff",
                        textShadow: "0 2px 4px rgba(0,0,0,0.3)",
                        zIndex: 10,
                        textAlign: "center",
                      }}
                    >
                      {student.nama}
                    </p>
                    <div
                      style={{
                        marginTop: "20px",
                        textAlign: "center",
                        padding: "0 20px",
                      }}
                    >
                      <p
                        style={{
                          fontSize: "10px",
                          fontWeight: "500",
                          opacity: 0.8,
                          lineHeight: "1.6",
                        }}
                      >
                        Simpan kartu ini dengan baik. Gunakan barcode di atas
                        untuk melakukan presensi pada mesin yang tersedia di
                        sekolah.
                      </p>
                      <p
                        style={{
                          fontSize: "9px",
                          fontWeight: "900",
                          marginTop: "20px",
                          opacity: 0.4,
                          letterSpacing: "0.2em",
                        }}
                      >
                        MTS NEGERI 2 BOMBANA
                      </p>
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </main>
    </div>
  );
}
