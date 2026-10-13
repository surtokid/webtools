/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useEffect, useRef, useState, useCallback } from 'react';
import { createClient } from '@supabase/supabase-js';
import L from 'leaflet';
import * as XLSX from 'xlsx';

// Fix Leaflet marker icon in Vite bundler
delete (L.Icon.Default.prototype as unknown as { _getIconUrl?: unknown })._getIconUrl;
L.Icon.Default.mergeOptions({
  iconRetinaUrl: 'https://unpkg.com/leaflet@1.9.4/dist/images/marker-icon-2x.png',
  iconUrl: 'https://unpkg.com/leaflet@1.9.4/dist/images/marker-icon.png',
  shadowUrl: 'https://unpkg.com/leaflet@1.9.4/dist/images/marker-shadow.png',
});

// Konfigurasi Kredensial Supabase & Google Apps Script
const SUPABASE_PROJECT_URL = "https://aofmqlbxnbcjriilubjz.supabase.co";
const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImFvZm1xbGJ4bmJjanJpaWx1Ymp6Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3OTE2MzEzMTIsImV4cCI6MjEwNzIwNzMxMn0.59wDqLyJZ9FH7q4Al6Do4XpDTFuP2DXm4HvQa4aGVN0";
const GAS_ENDPOINT_URL = "https://script.google.com/macros/s/AKfycbwsEYjlhaNqkZUGJVd2R4WHiaybiNPCe1rF6Y8wzopahQg09D9kqGnqsPVCu7Q2IKWl1g/exec";
const GDRIVE_FOLDER_ID = "10QkIZG7iJbdcHwRTyKZt1EKlqADJ8V36";
const REGION_API_BASE = "https://www.emsifa.com/api-wilayah-indonesia/api";

const supabaseClient = createClient(SUPABASE_PROJECT_URL, SUPABASE_ANON_KEY);

// Tipe Data
interface Karyawan {
  id: string;
  nama_lengkap: string;
  tgl_lahir?: string;
  username: string;
  no_wa: string;
  email: string;
  alamat?: string;
  divisi?: string;
  shifts?: string[];
  sisa_libur?: number;
  sisa_cuti?: number;
  password?: string;
}

interface AdminUser {
  id: string;
  nama_lengkap: string;
  username: string;
  email: string;
  role?: string;
  password?: string;
}

interface Lokasi {
  id: string;
  nama_lokasi: string;
  latitude: number;
  longitude: number;
  radius_meter: number;
  created_at?: string;
}

interface ShiftItem {
  id: string;
  nama_shift: string;
  jam_masuk: string;
  jam_pulang: string;
}

interface DivisiItem {
  id: string;
  nama_divisi: string;
  shifts?: string[];
  allow_cuti?: boolean;
}

interface PresensiRecord {
  id: string;
  karyawan_id: string;
  karyawan_nama: string;
  divisi?: string;
  tipe: 'masuk' | 'keluar';
  status?: string;
  shift_nama?: string;
  lokasi_id?: string;
  lokasi_nama?: string;
  latitude?: number;
  longitude?: number;
  foto_url?: string;
  keterangan?: string;
  tanggal: string;
  created_at: string;
}

interface LemburRecord {
  id: string;
  karyawan_id: string;
  karyawan_nama: string;
  durasi_jam: number;
  keterangan: string;
  bukti_url?: string;
  status: 'Pending' | 'Approved' | 'Rejected';
  tanggal: string;
  created_at?: string;
}

interface PengumumanRecord {
  id: string;
  judul: string;
  isi: string;
  author?: string;
  created_at: string;
}

interface GroupedRekap {
  karyawan_id: string;
  karyawan_nama: string;
  divisi: string;
  shift_nama: string;
  lokasi_nama: string;
  masuk: PresensiRecord | null;
  keluar: PresensiRecord | null;
  records: PresensiRecord[];
  category?: 'masuk' | 'belum_masuk' | 'sakit' | 'libur' | 'cuti';
}

interface TeamMemberStatus {
  id: string;
  nama: string;
  divisi: string;
  category: 'masuk' | 'belum_masuk' | 'izin' | 'libur' | 'cuti';
  statusText: string;
  statusColor: string;
}

interface RegionItem {
  id: string;
  name: string;
}

export default function App() {
  // Navigation / Hash view
  const [currentView, setCurrentView] = useState<'karyawan' | 'dashboard' | 'status-karyawan'>('karyawan');

  // Auth States
  const [currentUser, setCurrentUser] = useState<Karyawan | null>(null);
  const [currentAdmin, setCurrentAdmin] = useState<AdminUser | null>(null);

  // Clock
  const [currentTime, setCurrentTime] = useState('');
  const [currentDateStr, setCurrentDateStr] = useState('');

  // Active Month History
  const [userHistoryRecords, setUserHistoryRecords] = useState<PresensiRecord[]>([]);
  const [todayAttendanceRecords, setTodayAttendanceRecords] = useState<PresensiRecord[]>([]);
  const [periodLabel, setPeriodLabel] = useState('Periode Aktif');
  const [todayShiftInfo, setTodayShiftInfo] = useState<{ text: string; icon: string; className: string }>({
    text: 'Jam Kerja: Belum Absen Masuk',
    icon: 'fa-solid fa-clock text-indigo-200',
    className: 'mt-2.5 inline-flex items-center gap-1.5 px-3 py-1 rounded-xl bg-white/15 backdrop-blur-md border border-white/20 text-[11px] font-bold text-white shadow-sm transition-all',
  });

  // Team daily status view
  const [teamDailyData, setTeamDailyData] = useState<TeamMemberStatus[]>([]);
  const [activeTeamFilter, setActiveTeamFilter] = useState<'masuk' | 'belum_masuk' | 'izin' | 'libur' | 'cuti'>('masuk');
  const [teamLoading, setTeamLoading] = useState(false);

  // Admin state
  const [activeAdminTab, setActiveAdminTab] = useState<'rekap' | 'karyawan' | 'lokasi' | 'shift' | 'divisi' | 'lembur' | 'jatah' | 'pengumuman' | 'useradmin'>('rekap');
  const [isSidebarCollapsed, setIsSidebarCollapsed] = useState(false);
  const [filterRekapTanggal, setFilterRekapTanggal] = useState('');
  const [rekapList, setRekapList] = useState<GroupedRekap[]>([]);
  const [rekapStats, setRekapStats] = useState({ masuk: 0, belum_masuk: 0, sakit: 0, libur: 0, cuti: 0 });
  const [rekapFilterCategory, setRekapFilterCategory] = useState<'all' | 'masuk' | 'belum_masuk' | 'sakit' | 'libur' | 'cuti'>('all');
  const [modalExportExcelOpen, setModalExportExcelOpen] = useState(false);
  const [exportStartDate, setExportStartDate] = useState('');
  const [exportEndDate, setExportEndDate] = useState('');
  const [exportExcelLoading, setExportExcelLoading] = useState(false);
  const [adminKaryawanList, setAdminKaryawanList] = useState<Karyawan[]>([]);
  const [adminLokasiList, setAdminLokasiList] = useState<Lokasi[]>([]);
  const [adminShiftList, setAdminShiftList] = useState<ShiftItem[]>([]);
  const [toleransiMenit, setToleransiMenit] = useState<number>(15);
  const [adminDivisiList, setAdminDivisiList] = useState<DivisiItem[]>([]);
  const [adminLemburList, setAdminLemburList] = useState<LemburRecord[]>([]);
  const [adminPengumumanList, setAdminPengumumanList] = useState<PengumumanRecord[]>([]);
  const [adminUsersList, setAdminUsersList] = useState<AdminUser[]>([]);
  const [searchKaryawanQuery, setSearchKaryawanQuery] = useState('');
  const [searchRekapQuery, setSearchRekapQuery] = useState('');

  const filteredAdminKaryawanList = React.useMemo(() => {
    if (!searchKaryawanQuery.trim()) return adminKaryawanList;
    const q = searchKaryawanQuery.toLowerCase().trim();
    return adminKaryawanList.filter(k =>
      k.nama_lengkap.toLowerCase().includes(q) ||
      k.username.toLowerCase().includes(q) ||
      (k.divisi && k.divisi.toLowerCase().includes(q))
    );
  }, [adminKaryawanList, searchKaryawanQuery]);

  const filteredRekapList = React.useMemo(() => {
    let list = rekapList;
    if (rekapFilterCategory !== 'all') {
      list = list.filter(item => item.category === rekapFilterCategory);
    }
    if (searchRekapQuery.trim()) {
      const q = searchRekapQuery.toLowerCase().trim();
      list = list.filter(item =>
        item.karyawan_nama.toLowerCase().includes(q) ||
        item.divisi.toLowerCase().includes(q) ||
        item.shift_nama.toLowerCase().includes(q) ||
        (item.lokasi_nama && item.lokasi_nama.toLowerCase().includes(q))
      );
    }
    return list;
  }, [rekapList, rekapFilterCategory, searchRekapQuery]);

  // Modals visibility
  const [modalPresensiOpen, setModalPresensiOpen] = useState(false);
  const [presensiMode, setPresensiMode] = useState<'masuk' | 'keluar'>('masuk');
  const [presensiStep, setPresensiStep] = useState(1);
  const [cachedLokasi, setCachedLokasi] = useState<Lokasi[]>([]);
  const [selectedOfficeId, setSelectedOfficeId] = useState<string>('');
  const [selectedShiftName, setSelectedShiftName] = useState<string>('');
  const [statusKeluar, setStatusKeluar] = useState<'PULANG' | 'IZIN' | 'SAKIT' | 'LIBUR' | 'CUTI'>('PULANG');
  const [presensiKeterangan, setPresensiKeterangan] = useState('');
  const [userGps, setUserGps] = useState<{ lat: number; lng: number } | null>(null);
  const [gpsStatusText, setGpsStatusText] = useState('Mendeteksi GPS...');
  const [gpsBadgeClass, setGpsBadgeClass] = useState('text-[10px] font-bold px-2 py-0.5 rounded-full bg-slate-100 text-slate-600');
  const [isInsideGeofence, setIsInsideGeofence] = useState(false);
  const [capturedPhoto, setCapturedPhoto] = useState<string | null>(null);
  const [isPresensiSubmitting, setIsPresensiSubmitting] = useState(false);

  // Pengumuman state for employee view
  const [karyawanPengumumanList, setKaryawanPengumumanList] = useState<PengumumanRecord[]>([]);
  const [modalPengumumanDetail, setModalPengumumanDetail] = useState<PengumumanRecord | null>(null);
  const [modalAllPengumumanOpen, setModalAllPengumumanOpen] = useState(false);

  // Modal Lembur
  const [modalLemburOpen, setModalLemburOpen] = useState(false);
  const [lemburFoto, setLemburFoto] = useState<string | null>(null);
  const [lemburDurasi, setLemburDurasi] = useState<string>('');
  const [lemburKeterangan, setLemburKeterangan] = useState<string>('');
  const [lemburSubmitting, setLemburSubmitting] = useState(false);

  // Modal Izin / Sakit / Cuti
  const [modalIzinOpen, setModalIzinOpen] = useState(false);
  const [izinType, setIzinType] = useState<'SAKIT' | 'IZIN' | 'CUTI'>('SAKIT');
  const [izinFoto, setIzinFoto] = useState<string | null>(null);
  const [izinKeterangan, setIzinKeterangan] = useState('');
  const [izinSubmitting, setIzinSubmitting] = useState(false);

  // Modal Libur Direct
  const [modalLiburOpen, setModalLiburOpen] = useState(false);
  const [liburKeterangan, setLiburKeterangan] = useState('');
  const [liburSubmitting, setLiburSubmitting] = useState(false);

  // Modal Cuti Locked
  const [modalCutiLockedOpen, setModalCutiLockedOpen] = useState(false);

  // Modal PWA Guide
  const [modalPwaGuideOpen, setModalPwaGuideOpen] = useState(false);
  const [deferredPrompt, setDeferredPrompt] = useState<any>(null);

  // Modal Photo Preview
  const [photoPreviewModalOpen, setPhotoPreviewModalOpen] = useState(false);
  const [previewPhotoUrl, setPreviewPhotoUrl] = useState('');
  const [previewPhotoTitle, setPreviewPhotoTitle] = useState('');
  const [previewPhotoSubtitle, setPreviewPhotoSubtitle] = useState('');
  const [previewPhotoLoading, setPreviewPhotoLoading] = useState(true);
  const [previewPhotoError, setPreviewPhotoError] = useState(false);

  // Modal Edit Rekap
  const [modalEditRekapOpen, setModalEditRekapOpen] = useState(false);
  const [editingRekapItem, setEditingRekapItem] = useState<GroupedRekap | null>(null);
  const [editRekapShift, setEditRekapShift] = useState('');
  const [editRekapJamMasuk, setEditRekapJamMasuk] = useState('');
  const [editRekapStatusMasuk, setEditRekapStatusMasuk] = useState('Hadir');
  const [editRekapJamKeluar, setEditRekapJamKeluar] = useState('');
  const [editRekapStatusKeluar, setEditRekapStatusKeluar] = useState('PULANG');
  const [editRekapKeterangan, setEditRekapKeterangan] = useState('');
  const [editRekapSaving, setEditRekapSaving] = useState(false);

  // Modal Karyawan Admin Form
  const [modalKaryawanOpen, setModalKaryawanOpen] = useState(false);
  const [editingKaryawanId, setEditingKaryawanId] = useState<string | null>(null);
  const [karyawanFormNama, setKaryawanFormNama] = useState('');
  const [karyawanFormTglLahir, setKaryawanFormTglLahir] = useState('');
  const [karyawanFormUsername, setKaryawanFormUsername] = useState('');
  const [karyawanFormWa, setKaryawanFormWa] = useState('');
  const [karyawanFormEmail, setKaryawanFormEmail] = useState('');
  const [karyawanFormPass, setKaryawanFormPass] = useState('');
  const [karyawanFormDivisi, setKaryawanFormDivisi] = useState('');
  const [karyawanFormJalan, setKaryawanFormJalan] = useState('');
  const [provinsiList, setProvinsiList] = useState<RegionItem[]>([]);
  const [kotaList, setKotaList] = useState<RegionItem[]>([]);
  const [kecamatanList, setKecamatanList] = useState<RegionItem[]>([]);
  const [desaList, setDesaList] = useState<RegionItem[]>([]);
  const [selectedProvId, setSelectedProvId] = useState('');
  const [selectedKotaId, setSelectedKotaId] = useState('');
  const [selectedKecId, setSelectedKecId] = useState('');
  const [selectedDesaId, setSelectedDesaId] = useState('');
  const [regionLoading, setRegionLoading] = useState(false);

  // Modal Lokasi Form
  const [modalLokasiOpen, setModalLokasiOpen] = useState(false);
  const [editingLokasiId, setEditingLokasiId] = useState<string | null>(null);
  const [lokasiNama, setLokasiNama] = useState('');
  const [lokasiLat, setLokasiLat] = useState('');
  const [lokasiLng, setLokasiLng] = useState('');
  const [lokasiRadius, setLokasiRadius] = useState('100');

  // Modal Shift Form
  const [modalShiftOpen, setModalShiftOpen] = useState(false);
  const [editingShiftId, setEditingShiftId] = useState<string | null>(null);
  const [shiftNama, setShiftNama] = useState('');
  const [shiftJamMasuk, setShiftJamMasuk] = useState('');
  const [shiftJamPulang, setShiftJamPulang] = useState('');

  // Modal Divisi Form
  const [modalDivisiOpen, setModalDivisiOpen] = useState(false);
  const [editingDivisiId, setEditingDivisiId] = useState<string | null>(null);
  const [divisiNama, setDivisiNama] = useState('');
  const [divisiSelectedShifts, setDivisiSelectedShifts] = useState<string[]>([]);
  const [divisiAllowCuti, setDivisiAllowCuti] = useState(true);

  // Modal Admin Form
  const [modalAdminOpen, setModalAdminOpen] = useState(false);
  const [editingAdminId, setEditingAdminId] = useState<string | null>(null);
  const [adminUserNama, setAdminUserNama] = useState('');
  const [adminUserUsername, setAdminUserUsername] = useState('');
  const [adminUserEmail, setAdminUserEmail] = useState('');
  const [adminUserPass, setAdminUserPass] = useState('');

  // Modal Confirm Delete
  const [modalConfirmDeleteOpen, setModalConfirmDeleteOpen] = useState(false);
  const [deleteConfirmMsg, setDeleteConfirmMsg] = useState('Apakah Anda yakin ingin menghapus data ini secara permanen?');
  const [deleteActionFn, setDeleteActionFn] = useState<(() => Promise<void>) | null>(null);
  const [deleteExecuting, setDeleteExecuting] = useState(false);

  // Toast notifications
  const [toasts, setToasts] = useState<Array<{ id: number; message: string; type: 'info' | 'success' | 'error' | 'warning' }>>([]);

  // Refs for media & leaflet
  const videoRef = useRef<HTMLVideoElement | null>(null);
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const cameraStreamRef = useRef<MediaStream | null>(null);
  const leafletMapRef = useRef<L.Map | null>(null);
  const leafletMarkerRef = useRef<L.Marker | null>(null);
  const leafletCircleRef = useRef<L.Circle | null>(null);
  const mapContainerRef = useRef<HTMLDivElement | null>(null);

  // Polling tracker refs
  const lastKnownPresensiIdRef = useRef<string | null>(null);
  const lastKnownPengumumanIdRef = useRef<string | null>(null);

  // Helper Toast
  const showToast = useCallback((message: string, type: 'info' | 'success' | 'error' | 'warning' = 'info') => {
    const id = Date.now() + Math.random();
    setToasts(prev => [...prev, { id, message, type }]);
    setTimeout(() => {
      setToasts(prev => prev.filter(t => t.id !== id));
    }, 4000);
  }, []);

  // Audio notification chime
  const playDingSound = useCallback(() => {
    try {
      const AudioCtx = window.AudioContext || (window as unknown as { webkitAudioContext: typeof AudioContext }).webkitAudioContext;
      const audioCtx = new AudioCtx();
      const osc = audioCtx.createOscillator();
      const gain = audioCtx.createGain();
      osc.connect(gain);
      gain.connect(audioCtx.destination);
      osc.frequency.setValueAtTime(587.33, audioCtx.currentTime);
      osc.frequency.setValueAtTime(880, audioCtx.currentTime + 0.1);
      gain.gain.setValueAtTime(0.15, audioCtx.currentTime);
      gain.gain.exponentialRampToValueAtTime(0.001, audioCtx.currentTime + 0.5);
      osc.start();
      osc.stop(audioCtx.currentTime + 0.5);
    } catch {
      // Ignore if not permitted
    }
  }, []);

  const getLocalISODate = useCallback(() => {
    const d = new Date();
    d.setMinutes(d.getMinutes() - d.getTimezoneOffset());
    return d.toISOString().split('T')[0];
  }, []);

  // Upload photo to Google Drive via GAS
  const uploadPhotoToGoogleDrive = async (base64Data: string, fileName: string): Promise<string | null> => {
    try {
      const payload = {
        image: base64Data,
        fileName: fileName,
        folderId: GDRIVE_FOLDER_ID,
        karyawanName: currentUser?.nama_lengkap || 'Karyawan'
      };

      const response = await fetch(GAS_ENDPOINT_URL, {
        method: 'POST',
        headers: {
          'Content-Type': 'text/plain;charset=utf-8'
        },
        body: JSON.stringify(payload)
      });

      const resJson = await response.json();
      if (resJson && resJson.fileUrl) {
        return resJson.fileUrl;
      } else if (resJson && resJson.alternateViewUrl) {
        return resJson.alternateViewUrl;
      }
      return null;
    } catch (err) {
      console.warn('Upload GDrive via GAS terlewati:', err);
      return null;
    }
  };

  // Date utilities
  const formatIndoDate = (dateStr: string) => {
    if (!dateStr) return '';
    const parts = String(dateStr).split('-');
    if (parts.length !== 3) return dateStr;
    const y = parseInt(parts[0], 10);
    const m = parseInt(parts[1], 10) - 1;
    const d = parseInt(parts[2], 10);
    const bulanIndo = [
      'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
      'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
    ];
    return `${d} ${bulanIndo[m]} ${y}`;
  };

  const formatStatusWord = (str?: string) => {
    if (!str) return '';
    const s = String(str).trim();
    return s.charAt(0).toUpperCase() + s.slice(1).toLowerCase();
  };

  const formatTime24 = (dateStr?: string) => {
    if (!dateStr) return '';
    const d = new Date(dateStr);
    return d.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' }).replace('.', ':');
  };

  const ddmmyyyyToIso = (val: string) => {
    if (!val) return null;
    const p = val.trim().split('/');
    if (p.length === 3 && p[0] && p[1] && p[2]) {
      return `${p[2]}-${p[1].padStart(2, '0')}-${p[0].padStart(2, '0')}`;
    }
    return val;
  };

  const isoToDdmmyyyy = (val?: string) => {
    if (!val) return '';
    const p = String(val).split('-');
    if (p.length === 3) {
      return `${p[2].padStart(2, '0')}/${p[1].padStart(2, '0')}/${p[0]}`;
    }
    return val;
  };

  const parseAddressData = (rawAlamat?: string | null) => {
    if (!rawAlamat) return { jalan: '', provId: '', provName: '', kotaId: '', kotaName: '', kecId: '', kecName: '', desaId: '', desaName: '', fullAddress: '' };
    const trimmed = String(rawAlamat).trim();
    if (trimmed.startsWith('{') && trimmed.endsWith('}')) {
      try {
        const obj = JSON.parse(trimmed);
        return {
          jalan: obj.jalan || '',
          provId: obj.provId || '',
          provName: obj.provName || '',
          kotaId: obj.kotaId || '',
          kotaName: obj.kotaName || '',
          kecId: obj.kecId || '',
          kecName: obj.kecName || '',
          desaId: obj.desaId || '',
          desaName: obj.desaName || '',
          fullAddress: obj.fullAddress || ''
        };
      } catch {}
    }
    return {
      jalan: trimmed,
      provId: '',
      provName: '',
      kotaId: '',
      kotaName: '',
      kecId: '',
      kecName: '',
      desaId: '',
      desaName: '',
      fullAddress: trimmed
    };
  };

  const getDisplayAddress = (rawAlamat?: string | null) => {
    if (!rawAlamat) return '-';
    const parsed = parseAddressData(rawAlamat);
    if (parsed.fullAddress) return parsed.fullAddress;
    return parsed.jalan || '-';
  };

  const formatAutoDateInput = (val: string) => {
    let clean = val.replace(/\D/g, '');
    if (clean.length > 8) clean = clean.substring(0, 8);

    if (clean.length >= 5) {
      let d = clean.substring(0, 2);
      let m = clean.substring(2, 4);
      const y = clean.substring(4);
      if (parseInt(d) > 31) d = '31';
      if (parseInt(m) > 12) m = '12';
      return d + '/' + m + '/' + y;
    } else if (clean.length >= 3) {
      let d = clean.substring(0, 2);
      let m = clean.substring(2);
      if (parseInt(d) > 31) d = '31';
      if (m.length === 2 && parseInt(m) > 12) m = '12';
      return d + '/' + m;
    } else if (clean.length >= 2) {
      let d = clean.substring(0, 2);
      if (parseInt(d) > 31) d = '31';
      return d + '/';
    }
    return clean;
  };

  const formatAuto24HourTime = (val: string) => {
    let clean = val.replace(/\D/g, '');
    if (clean.length > 4) clean = clean.substring(0, 4);

    if (clean.length >= 3) {
      let h = parseInt(clean.substring(0, 2));
      const m = clean.substring(2);
      if (h > 23) h = 23;
      return String(h).padStart(2, '0') + ':' + m;
    } else if (clean.length >= 2) {
      let h = parseInt(clean.substring(0, 2));
      if (h > 23) h = 23;
      return String(h).padStart(2, '0') + ':';
    }
    return clean;
  };

  const getActivePeriod = () => {
    const now = new Date();
    const y = now.getFullYear();
    const m = now.getMonth();
    const d = now.getDate();

    let startDate: Date, endDate: Date;
    if (d <= 25) {
      startDate = new Date(y, m - 1, 26);
      endDate = new Date(y, m, 25);
    } else {
      startDate = new Date(y, m, 26);
      endDate = new Date(y, m + 1, 25);
    }

    const bulanIndo = [
      'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
      'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
    ];

    const pad = (n: number) => String(n).padStart(2, '0');
    const startIso = `${startDate.getFullYear()}-${pad(startDate.getMonth() + 1)}-${pad(startDate.getDate())}`;
    const endIso = `${endDate.getFullYear()}-${pad(endDate.getMonth() + 1)}-${pad(endDate.getDate())}`;

    let label: string;
    if (startDate.getFullYear() === endDate.getFullYear()) {
      label = `${startDate.getDate()} ${bulanIndo[startDate.getMonth()]} - ${endDate.getDate()} ${bulanIndo[endDate.getMonth()]} ${endDate.getFullYear()}`;
    } else {
      label = `${startDate.getDate()} ${bulanIndo[startDate.getMonth()]} ${startDate.getFullYear()} - ${endDate.getDate()} ${bulanIndo[endDate.getMonth()]} ${endDate.getFullYear()}`;
    }

    return { startIso, endIso, label };
  };

  const formatDailyActivityText = (records: PresensiRecord[]) => {
    if (!records || records.length === 0) {
      return { text: '-', type: 'neutral' };
    }

    const masukRec = records.find(r => r.tipe === 'masuk');
    const keluarRec = records.find(r => r.tipe === 'keluar');

    if (!masukRec) {
      const liburRec = records.find(r => (r.status || '').toUpperCase() === 'LIBUR');
      if (liburRec) return { text: 'Libur', type: 'libur' };

      const cutiRec = records.find(r => (r.status || '').toUpperCase() === 'CUTI');
      if (cutiRec) return { text: 'Cuti', type: 'cuti' };

      const sakitRec = records.find(r => (r.status || '').toUpperCase() === 'SAKIT');
      if (sakitRec) {
        const tSakit = formatTime24(sakitRec.created_at);
        return { text: `Sakit${tSakit ? ` (${tSakit})` : ''}`, type: 'sakit' };
      }

      const izinRec = records.find(r => (r.status || '').toUpperCase() === 'IZIN');
      if (izinRec) {
        const tIzin = formatTime24(izinRec.created_at);
        return { text: `Izin${tIzin ? ` (${tIzin})` : ''}`, type: 'izin' };
      }

      if (keluarRec) {
        const st = formatStatusWord(keluarRec.status || 'Pulang');
        const tKeluar = formatTime24(keluarRec.created_at);
        const shift = keluarRec.shift_nama ? ` • ${keluarRec.shift_nama}` : '';
        return { text: `Belum Masuk - ${st} (${tKeluar})${shift}`, type: 'keluar' };
      }

      return { text: '-', type: 'neutral' };
    }

    let stMasuk = 'Masuk';
    const rawMasukStatus = (masukRec.status || '').toUpperCase();
    if (rawMasukStatus === 'TERLAMBAT') {
      stMasuk = 'Terlambat';
    } else if (rawMasukStatus === 'HADIR' || rawMasukStatus === 'MASUK') {
      stMasuk = 'Masuk';
    } else if (masukRec.status) {
      stMasuk = formatStatusWord(masukRec.status);
    }
    const timeMasuk = formatTime24(masukRec.created_at);
    const masukPart = timeMasuk ? `${stMasuk} (${timeMasuk})` : stMasuk;

    let stPulang = 'Belum Pulang';
    let type = 'masuk';

    if (keluarRec) {
      const rawKeluarStatus = (keluarRec.status || '').toUpperCase();
      if (rawKeluarStatus === 'PULANG') {
        stPulang = 'Pulang';
        type = 'keluar';
      } else if (rawKeluarStatus === 'LIBUR') {
        stPulang = 'Libur';
        type = 'libur';
      } else if (rawKeluarStatus === 'CUTI') {
        stPulang = 'Cuti';
        type = 'cuti';
      } else if (rawKeluarStatus === 'SAKIT') {
        stPulang = 'Sakit';
        type = 'sakit';
      } else if (rawKeluarStatus === 'IZIN') {
        stPulang = 'Izin';
        type = 'izin';
      } else {
        stPulang = formatStatusWord(keluarRec.status);
        type = 'keluar';
      }

      const timeKeluar = formatTime24(keluarRec.created_at);
      stPulang = timeKeluar ? `${stPulang} (${timeKeluar})` : stPulang;
    }

    const shiftNama = masukRec.shift_nama || keluarRec?.shift_nama || '';
    const shiftPart = shiftNama ? ` • ${shiftNama}` : '';

    return { text: `${masukPart} - ${stPulang}${shiftPart}`, type: type };
  };

  const calculateDistanceMeters = (lat1: number, lon1: number, lat2: number, lon2: number) => {
    const R = 6371e3;
    const φ1 = (lat1 * Math.PI) / 180;
    const φ2 = (lat2 * Math.PI) / 180;
    const Δφ = ((lat2 - lat1) * Math.PI) / 180;
    const Δλ = ((lon2 - lon1) * Math.PI) / 180;
    const a = Math.sin(Δφ / 2) * Math.sin(Δφ / 2) + Math.cos(φ1) * Math.cos(φ2) * Math.sin(Δλ / 2) * Math.sin(Δλ / 2);
    const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
    return Math.round(R * c);
  };

  const formatDirectImageUrl = (url: string) => {
    if (!url) return '';
    if (url.startsWith('data:image')) return url;
    if (url.includes('lh3.googleusercontent.com')) return url;
    const matchFile = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
    if (matchFile && matchFile[1]) {
      return `https://lh3.googleusercontent.com/d/${matchFile[1]}`;
    }
    const matchId = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    if (matchId && matchId[1]) {
      return `https://lh3.googleusercontent.com/d/${matchId[1]}`;
    }
    return url;
  };

  // Hash Navigation
  const navigateTo = (hash: string) => {
    window.location.hash = hash;
  };

  useEffect(() => {
    const handleHash = () => {
      const hash = window.location.hash.toLowerCase();
      if (hash === '#dashboard') {
        setCurrentView('dashboard');
      } else if (hash === '#status-karyawan') {
        setCurrentView('status-karyawan');
      } else {
        setCurrentView('karyawan');
      }
    };
    handleHash();
    window.addEventListener('hashchange', handleHash);
    return () => window.removeEventListener('hashchange', handleHash);
  }, []);

  // Clock interval
  useEffect(() => {
    const timer = setInterval(() => {
      const now = new Date();
      setCurrentTime(now.toLocaleTimeString('id-ID', { hour12: false }));
      setCurrentDateStr(now.toLocaleDateString('id-ID', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' }));
    }, 1000);
    return () => clearInterval(timer);
  }, []);

  // Check saved session on load
  useEffect(() => {
    const u = localStorage.getItem('presensi_user');
    const a = localStorage.getItem('presensi_admin');
    if (u) {
      try {
        setCurrentUser(JSON.parse(u));
      } catch {}
    }
    if (a) {
      try {
        setCurrentAdmin(JSON.parse(a));
      } catch {}
    }
    const col = localStorage.getItem('presensi_admin_sidebar_collapsed');
    if (col === 'true') {
      setIsSidebarCollapsed(true);
    }
    setFilterRekapTanggal(getLocalISODate());
  }, [getLocalISODate]);

  // Muat pengumuman untuk tampilan karyawan
  const loadKaryawanPengumuman = useCallback(async () => {
    try {
      const { data } = await supabaseClient
        .from('pengumuman_presensi')
        .select('*')
        .order('created_at', { ascending: false })
        .limit(20);
      setKaryawanPengumumanList(data || []);
    } catch (err) {
      console.warn('Gagal memuat pengumuman karyawan:', err);
    }
  }, []);

  // Load user data & history when logged in
  const refreshUserDashboard = useCallback(async () => {
    if (!currentUser) return;
    const { data: refreshedUser } = await supabaseClient
      .from('karyawan_presensi')
      .select('sisa_libur, sisa_cuti')
      .eq('id', currentUser.id)
      .single();

    if (refreshedUser) {
      setCurrentUser(prev => prev ? { ...prev, sisa_libur: refreshedUser.sisa_libur, sisa_cuti: refreshedUser.sisa_cuti } : null);
    }

    const period = getActivePeriod();
    setPeriodLabel(period.label);

    const { data: recs } = await supabaseClient
      .from('presensi_presensi')
      .select('*')
      .eq('karyawan_id', currentUser.id)
      .gte('tanggal', period.startIso)
      .lte('tanggal', period.endIso)
      .order('tanggal', { ascending: false })
      .order('created_at', { ascending: false });

    const records = recs || [];
    setUserHistoryRecords(records);

    const today = getLocalISODate();
    const todayRecs = records.filter(r => r.tanggal === today);
    setTodayAttendanceRecords(todayRecs);

    // Update shift badge
    const masukRecord = todayRecs.find(r => r.tipe === 'masuk');
    const liburRecord = todayRecs.find(r => (r.status || '').toUpperCase() === 'LIBUR');
    const cutiRecord = todayRecs.find(r => (r.status || '').toUpperCase() === 'CUTI');
    const sakitRecord = todayRecs.find(r => (r.status || '').toUpperCase() === 'SAKIT');
    const izinRecord = todayRecs.find(r => (r.status || '').toUpperCase() === 'IZIN');

    if (masukRecord) {
      const shiftNama = masukRecord.shift_nama || 'Reguler';
      let jamKerjaStr = '';
      try {
        const { data: sData } = await supabaseClient
          .from('shift_presensi')
          .select('nama_shift, jam_masuk, jam_pulang')
          .eq('nama_shift', shiftNama)
          .maybeSingle();
        if (sData && sData.jam_masuk && sData.jam_pulang) {
          jamKerjaStr = `${sData.jam_masuk.substring(0, 5)} - ${sData.jam_pulang.substring(0, 5)} WIB`;
        }
      } catch {}

      setTodayShiftInfo({
        text: jamKerjaStr ? `Jam Kerja: ${jamKerjaStr} (${shiftNama})` : `Jam Kerja: ${shiftNama}`,
        icon: 'fa-solid fa-business-time text-emerald-300',
        className: 'mt-2.5 inline-flex items-center gap-1.5 px-3 py-1 rounded-xl bg-emerald-500/25 backdrop-blur-md border border-emerald-400/30 text-[11px] font-bold text-white shadow-sm transition-all',
      });
    } else if (liburRecord) {
      setTodayShiftInfo({
        text: 'Status: Libur Hari Ini',
        icon: 'fa-solid fa-umbrella-beach text-sky-300',
        className: 'mt-2.5 inline-flex items-center gap-1.5 px-3 py-1 rounded-xl bg-sky-500/25 backdrop-blur-md border border-sky-400/30 text-[11px] font-bold text-white shadow-sm transition-all',
      });
    } else if (cutiRecord) {
      setTodayShiftInfo({
        text: 'Status: Cuti',
        icon: 'fa-solid fa-plane-departure text-purple-300',
        className: 'mt-2.5 inline-flex items-center gap-1.5 px-3 py-1 rounded-xl bg-purple-500/25 backdrop-blur-md border border-purple-400/30 text-[11px] font-bold text-white shadow-sm transition-all',
      });
    } else if (sakitRecord) {
      setTodayShiftInfo({
        text: 'Status: Sakit',
        icon: 'fa-solid fa-briefcase-medical text-amber-300',
        className: 'mt-2.5 inline-flex items-center gap-1.5 px-3 py-1 rounded-xl bg-amber-500/25 backdrop-blur-md border border-amber-400/30 text-[11px] font-bold text-white shadow-sm transition-all',
      });
    } else if (izinRecord) {
      setTodayShiftInfo({
        text: 'Status: Izin',
        icon: 'fa-solid fa-envelope-open-text text-orange-300',
        className: 'mt-2.5 inline-flex items-center gap-1.5 px-3 py-1 rounded-xl bg-orange-500/25 backdrop-blur-md border border-orange-400/30 text-[11px] font-bold text-white shadow-sm transition-all',
      });
    } else {
      setTodayShiftInfo({
        text: 'Jam Kerja: Belum Absen Masuk',
        icon: 'fa-solid fa-clock text-indigo-200',
        className: 'mt-2.5 inline-flex items-center gap-1.5 px-3 py-1 rounded-xl bg-white/15 backdrop-blur-md border border-white/20 text-[11px] font-bold text-white shadow-sm transition-all',
      });
    }
  }, [currentUser, getLocalISODate]);

  useEffect(() => {
    loadKaryawanPengumuman();
    if (currentUser) {
      refreshUserDashboard();
    }
  }, [currentUser, refreshUserDashboard, loadKaryawanPengumuman]);

  // Load team daily status
  const loadEmployeeDailyStatus = useCallback(async () => {
    setTeamLoading(true);
    try {
      const today = getLocalISODate();
      const { data: emps, error: errEmp } = await supabaseClient
        .from('karyawan_presensi')
        .select('id, nama_lengkap, divisi')
        .order('nama_lengkap', { ascending: true });

      if (errEmp) throw errEmp;

      const { data: atts, error: errAtt } = await supabaseClient
        .from('presensi_presensi')
        .select('karyawan_id, tipe, status, created_at')
        .eq('tanggal', today)
        .order('created_at', { ascending: true });

      if (errAtt) throw errAtt;

      const teamData: TeamMemberStatus[] = (emps || []).map(emp => {
        const myAtts = (atts || []).filter(a => a.karyawan_id === emp.id);
        const masukRecord = myAtts.find(a => a.tipe === 'masuk');
        const keluarRecord = myAtts.find(a => a.tipe === 'keluar');
        const cutiRecord = myAtts.find(a => (a.status || '').toUpperCase() === 'CUTI');
        const liburRecord = myAtts.find(a => (a.status || '').toUpperCase() === 'LIBUR');
        const izinRecord = myAtts.find(a => ['IZIN', 'SAKIT'].includes((a.status || '').toUpperCase()));

        let category: 'masuk' | 'belum_masuk' | 'izin' | 'libur' | 'cuti' = 'belum_masuk';
        let statusText = 'Belum Masuk';
        let statusColor = 'bg-slate-100 text-slate-600 border-slate-200';

        if (cutiRecord) {
          category = 'cuti';
          statusText = 'Cuti';
          statusColor = 'bg-purple-100 text-purple-700 border-purple-200';
        } else if (liburRecord) {
          category = 'libur';
          statusText = 'Libur';
          statusColor = 'bg-sky-100 text-sky-700 border-sky-200';
        } else if (izinRecord && !masukRecord) {
          category = 'izin';
          statusText = izinRecord.status.toUpperCase() === 'SAKIT' ? 'Sakit' : 'Izin';
          statusColor = 'bg-amber-100 text-amber-700 border-amber-200';
        } else if (masukRecord) {
          category = 'masuk';
          if (keluarRecord) {
            statusText = 'Masuk - pulang';
            statusColor = 'bg-indigo-100 text-indigo-700 border-indigo-200';
          } else {
            statusText = 'Masuk - masih kerja';
            statusColor = 'bg-emerald-100 text-emerald-700 border-emerald-200';
          }
        }

        return {
          id: emp.id,
          nama: emp.nama_lengkap,
          divisi: emp.divisi || 'Umum',
          category,
          statusText,
          statusColor,
        };
      });

      setTeamDailyData(teamData);
    } catch (err) {
      console.warn('Gagal memuat status karyawan:', err);
    } finally {
      setTeamLoading(false);
    }
  }, [getLocalISODate]);

  useEffect(() => {
    if (currentView === 'status-karyawan') {
      loadEmployeeDailyStatus();
    }
  }, [currentView, loadEmployeeDailyStatus]);

  // Polling for live notifications
  useEffect(() => {
    const timer = setInterval(async () => {
      try {
        const { data: latestPresensi } = await supabaseClient
          .from('presensi_presensi')
          .select('id, karyawan_nama, shift_nama, tipe')
          .order('created_at', { ascending: false })
          .limit(1);

        if (latestPresensi && latestPresensi.length > 0) {
          const p = latestPresensi[0];
          if (lastKnownPresensiIdRef.current && p.id !== lastKnownPresensiIdRef.current) {
            playDingSound();
            showToast(`📢 ${p.karyawan_nama} telah ${p.tipe.toUpperCase()} (${p.shift_nama || 'Reguler'})`, 'info');
          }
          lastKnownPresensiIdRef.current = p.id;
        }

        const { data: latestPengumuman } = await supabaseClient
          .from('pengumuman_presensi')
          .select('id, judul, isi')
          .order('created_at', { ascending: false })
          .limit(1);

        if (latestPengumuman && latestPengumuman.length > 0) {
          const ann = latestPengumuman[0];
          if (lastKnownPengumumanIdRef.current && ann.id !== lastKnownPengumumanIdRef.current) {
            playDingSound();
            showToast(`🔔 PENGUMUMAN BARU: ${ann.judul}`, 'warning');
            loadKaryawanPengumuman();
          }
          lastKnownPengumumanIdRef.current = ann.id;
        }
      } catch {}
    }, 12000);
    return () => clearInterval(timer);
  }, [playDingSound, showToast]);

  // PWA Prompt listener
  useEffect(() => {
    const handleBeforeInstall = (e: any) => {
      e.preventDefault();
      setDeferredPrompt(e);
    };
    window.addEventListener('beforeinstallprompt', handleBeforeInstall);
    return () => window.removeEventListener('beforeinstallprompt', handleBeforeInstall);
  }, []);

  const triggerAddToHomeScreen = () => {
    if (deferredPrompt) {
      deferredPrompt.prompt();
      deferredPrompt.userChoice.then((choiceResult: any) => {
        if (choiceResult.outcome === 'accepted') {
          showToast('Memasang shortcut ke layar utama...', 'success');
        }
        setDeferredPrompt(null);
      });
    } else {
      setModalPwaGuideOpen(true);
    }
  };

  // Indonesian Region API cascade
  const fetchProvinces = async () => {
    setRegionLoading(true);
    try {
      const res = await fetch(`${REGION_API_BASE}/provinces.json`);
      const data = await res.json();
      setProvinsiList(data || []);
    } catch {
      showToast('Gagal memuat daftar provinsi.', 'error');
    } finally {
      setRegionLoading(false);
    }
  };

  const handleProvinsiChange = async (provId: string) => {
    setSelectedProvId(provId);
    setSelectedKotaId('');
    setSelectedKecId('');
    setSelectedDesaId('');
    setKotaList([]);
    setKecamatanList([]);
    setDesaList([]);
    if (!provId) return;

    setRegionLoading(true);
    try {
      const res = await fetch(`${REGION_API_BASE}/regencies/${provId}.json`);
      const data = await res.json();
      setKotaList(data || []);
    } catch {
      showToast('Gagal memuat kota.', 'error');
    } finally {
      setRegionLoading(false);
    }
  };

  const handleKotaChange = async (kotaId: string) => {
    setSelectedKotaId(kotaId);
    setSelectedKecId('');
    setSelectedDesaId('');
    setKecamatanList([]);
    setDesaList([]);
    if (!kotaId) return;

    setRegionLoading(true);
    try {
      const res = await fetch(`${REGION_API_BASE}/districts/${kotaId}.json`);
      const data = await res.json();
      setKecamatanList(data || []);
    } catch {
      showToast('Gagal memuat kecamatan.', 'error');
    } finally {
      setRegionLoading(false);
    }
  };

  const handleKecamatanChange = async (kecId: string) => {
    setSelectedKecId(kecId);
    setSelectedDesaId('');
    setDesaList([]);
    if (!kecId) return;

    setRegionLoading(true);
    try {
      const res = await fetch(`${REGION_API_BASE}/villages/${kecId}.json`);
      const data = await res.json();
      setDesaList(data || []);
    } catch {
      showToast('Gagal memuat desa/kelurahan.', 'error');
    } finally {
      setRegionLoading(false);
    }
  };


  // Buka dan pulihkan cascade wilayah karyawan secara otomatis
  const openKaryawanRegionCascade = async (addr: ReturnType<typeof parseAddressData>) => {
    setRegionLoading(true);
    try {
      // 1. Muat provinsi
      const resProv = await fetch(`${REGION_API_BASE}/provinces.json`);
      const dataProv = await resProv.json();
      setProvinsiList(dataProv || []);

      if (addr.provId) {
        setSelectedProvId(addr.provId);

        // 2. Muat Kota / Regencies
        const resKota = await fetch(`${REGION_API_BASE}/regencies/${addr.provId}.json`);
        const dataKota = await resKota.json();
        setKotaList(dataKota || []);

        if (addr.kotaId) {
          setSelectedKotaId(addr.kotaId);

          // 3. Muat Kecamatan / Districts
          const resKec = await fetch(`${REGION_API_BASE}/districts/${addr.kotaId}.json`);
          const dataKec = await resKec.json();
          setKecamatanList(dataKec || []);

          if (addr.kecId) {
            setSelectedKecId(addr.kecId);

            // 4. Muat Desa / Villages
            const resDesa = await fetch(`${REGION_API_BASE}/villages/${addr.kecId}.json`);
            const dataDesa = await resDesa.json();
            setDesaList(dataDesa || []);

            if (addr.desaId) {
              setSelectedDesaId(addr.desaId);
            } else {
              setSelectedDesaId('');
            }
          } else {
            setSelectedKecId('');
            setSelectedDesaId('');
            setDesaList([]);
          }
        } else {
          setSelectedKotaId('');
          setSelectedKecId('');
          setSelectedDesaId('');
          setKecamatanList([]);
          setDesaList([]);
        }
      } else {
        setSelectedProvId('');
        setSelectedKotaId('');
        setSelectedKecId('');
        setSelectedDesaId('');
        setKotaList([]);
        setKecamatanList([]);
        setDesaList([]);
      }
    } catch (err) {
      console.warn('Gagal memuat wilayah cascade:', err);
    } finally {
      setRegionLoading(false);
    }
  };

  // Admin Data loaders
  const loadAdminRekapan = useCallback(async (selectedDate?: string) => {
    const tgl = selectedDate || filterRekapTanggal || getLocalISODate();

    // 1. Ambil seluruh data master karyawan aktif
    const { data: allEmployees, error: errEmp } = await supabaseClient
      .from('karyawan_presensi')
      .select('id, nama_lengkap, divisi, shifts')
      .order('nama_lengkap', { ascending: true });

    if (errEmp) {
      console.warn('Gagal memuat master karyawan untuk rekap:', errEmp);
    }

    // 2. Ambil catatan presensi pada tanggal terkait
    const { data, error } = await supabaseClient
      .from('presensi_presensi')
      .select('*')
      .eq('tanggal', tgl)
      .order('created_at', { ascending: true });

    if (error) {
      showToast('Gagal memuat rekapan presensi.', 'error');
      return;
    }

    const mapKaryawan = new Map<string, GroupedRekap>();

    // Masukkan data presensi yang sudah tercatat
    (data || []).forEach(r => {
      const key = r.karyawan_id || r.karyawan_nama;
      if (!mapKaryawan.has(key)) {
        mapKaryawan.set(key, {
          karyawan_id: r.karyawan_id,
          karyawan_nama: r.karyawan_nama,
          divisi: r.divisi || '-',
          shift_nama: r.shift_nama || '-',
          lokasi_nama: r.lokasi_nama || '-',
          masuk: null,
          keluar: null,
          records: [],
          category: 'masuk'
        });
      }
      const item = mapKaryawan.get(key)!;
      item.records.push(r);

      if (r.tipe === 'masuk') {
        if (!item.masuk) item.masuk = r;
      } else {
        if (!item.keluar) item.keluar = r;
      }

      if (r.shift_nama && item.shift_nama === '-') item.shift_nama = r.shift_nama;
      if (r.divisi && item.divisi === '-') item.divisi = r.divisi;
      if (r.lokasi_nama && item.lokasi_nama === '-') item.lokasi_nama = r.lokasi_nama;
    });

    // Masukkan karyawan yang belum ada presensi di hari tersebut sebagai 'belum_masuk'
    (allEmployees || []).forEach(emp => {
      const key = emp.id;
      if (!mapKaryawan.has(key)) {
        mapKaryawan.set(key, {
          karyawan_id: emp.id,
          karyawan_nama: emp.nama_lengkap,
          divisi: emp.divisi || '-',
          shift_nama: (emp.shifts && emp.shifts[0]) || '-',
          lokasi_nama: '-',
          masuk: null,
          keluar: null,
          records: [],
          category: 'belum_masuk'
        });
      }
    });

    let cMasuk = 0, cBelumMasuk = 0, cSakit = 0, cLibur = 0, cCuti = 0;
    const grouped = Array.from(mapKaryawan.values()).map(item => {
      const stKeluar = (item.keluar?.status || '').toUpperCase();
      const stMasuk = (item.masuk?.status || '').toUpperCase();

      let cat: 'masuk' | 'belum_masuk' | 'sakit' | 'libur' | 'cuti' = 'belum_masuk';

      if (stKeluar === 'SAKIT' || stMasuk === 'SAKIT') {
        cat = 'sakit';
        cSakit++;
      } else if (stKeluar === 'LIBUR' || stMasuk === 'LIBUR') {
        cat = 'libur';
        cLibur++;
      } else if (stKeluar === 'CUTI' || stMasuk === 'CUTI') {
        cat = 'cuti';
        cCuti++;
      } else if (item.masuk || stKeluar === 'PULANG' || stKeluar === 'HADIR') {
        cat = 'masuk';
        cMasuk++;
      } else {
        cat = 'belum_masuk';
        cBelumMasuk++;
      }

      item.category = cat;
      return item;
    });

    setRekapList(grouped);
    setRekapStats({ masuk: cMasuk, belum_masuk: cBelumMasuk, sakit: cSakit, libur: cLibur, cuti: cCuti });
  }, [filterRekapTanggal, getLocalISODate, showToast]);

  const loadAdminKaryawan = async () => {
    const { data } = await supabaseClient.from('karyawan_presensi').select('*').order('nama_lengkap');
    setAdminKaryawanList(data || []);
  };

  const loadAdminLokasi = async () => {
    const { data } = await supabaseClient.from('lokasi_presensi').select('*').order('created_at');
    setAdminLokasiList(data || []);
  };

  const loadAdminShiftAndJam = async () => {
    const { data: shifts } = await supabaseClient.from('shift_presensi').select('*').order('jam_masuk');
    setAdminShiftList(shifts || []);
    const { data: tol } = await supabaseClient.from('jam_kerja_presensi').select('toleransi_menit').limit(1).single();
    if (tol) setToleransiMenit(tol.toleransi_menit);
  };

  const loadAdminDivisi = async () => {
    const { data } = await supabaseClient.from('divisi_presensi').select('*').order('nama_divisi', { ascending: true });
    setAdminDivisiList(data || []);
  };

  const loadAdminLembur = async () => {
    const { data } = await supabaseClient.from('lembur_presensi').select('*').order('created_at', { ascending: false });
    setAdminLemburList(data || []);
  };

  const loadAdminPengumuman = async () => {
    const { data } = await supabaseClient.from('pengumuman_presensi').select('*').order('created_at', { ascending: false });
    setAdminPengumumanList(data || []);
  };

  const loadAdminUsers = async () => {
    const { data } = await supabaseClient.from('admin_presensi').select('*').order('created_at');
    setAdminUsersList(data || []);
  };

  const switchAdminTab = (tab: typeof activeAdminTab) => {
    setActiveAdminTab(tab);
    if (tab === 'rekap') loadAdminRekapan();
    else if (tab === 'karyawan') loadAdminKaryawan();
    else if (tab === 'lokasi') loadAdminLokasi();
    else if (tab === 'shift') loadAdminShiftAndJam();
    else if (tab === 'divisi') loadAdminDivisi();
    else if (tab === 'lembur') loadAdminLembur();
    else if (tab === 'jatah') loadAdminKaryawan();
    else if (tab === 'pengumuman') loadAdminPengumuman();
    else if (tab === 'useradmin') loadAdminUsers();
  };

  useEffect(() => {
    if (currentView === 'dashboard' && currentAdmin) {
      loadAdminRekapan();
    }
  }, [currentView, currentAdmin, loadAdminRekapan]);

  // Modal Presensi camera and Leaflet map lifecycle
  const stopLiveCamera = () => {
    if (cameraStreamRef.current) {
      cameraStreamRef.current.getTracks().forEach(t => t.stop());
      cameraStreamRef.current = null;
    }
  };

  const startLiveCamera = async () => {
    setCapturedPhoto(null);
    try {
      const stream = await navigator.mediaDevices.getUserMedia({
        video: { facingMode: 'user', width: { ideal: 640 }, height: { ideal: 640 } },
        audio: false
      });
      cameraStreamRef.current = stream;
      if (videoRef.current) {
        videoRef.current.srcObject = stream;
      }
    } catch {
      showToast('Kamera tidak dapat diakses. Mohon beri izin kamera pada browser.', 'error');
    }
  };

  const takeSnapshot = () => {
    if (!videoRef.current || !canvasRef.current) return;
    const video = videoRef.current;
    const canvas = canvasRef.current;
    canvas.width = video.videoWidth || 480;
    canvas.height = video.videoHeight || 480;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    ctx.translate(canvas.width, 0);
    ctx.scale(-1, 1);
    ctx.drawImage(video, 0, 0, canvas.width, canvas.height);
    const photo = canvas.toDataURL('image/jpeg', 0.85);
    setCapturedPhoto(photo);
    stopLiveCamera();
  };

  const retakeSnapshot = () => {
    setCapturedPhoto(null);
    startLiveCamera();
  };

  const openModalPresensi = async (mode: 'masuk' | 'keluar') => {
    if (!currentUser) return;
    setPresensiMode(mode);
    setPresensiStep(1);
    setCapturedPhoto(null);
    setPresensiKeterangan('');
    setStatusKeluar('PULANG');

    // Load locations
    const { data: locs } = await supabaseClient.from('lokasi_presensi').select('*');
    const availableLocs = locs || [];
    setCachedLokasi(availableLocs);
    if (availableLocs.length > 0) {
      setSelectedOfficeId(availableLocs[0].id);
    }

    // Set shifts
    const userShifts = currentUser.shifts && currentUser.shifts.length > 0 ? currentUser.shifts : ['Reguler'];
    setSelectedShiftName(userShifts[0]);

    setModalPresensiOpen(true);
  };

  // Leaflet map initialization when Step 1 is rendered
  useEffect(() => {
    if (modalPresensiOpen && presensiStep === 1) {
      setTimeout(() => {
        if (!mapContainerRef.current) return;

        if (!leafletMapRef.current) {
          const map = L.map(mapContainerRef.current).setView([-6.9932, 110.4203], 15);
          L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', { maxZoom: 19 }).addTo(map);
          leafletMapRef.current = map;
        } else {
          leafletMapRef.current.invalidateSize();
        }

        navigator.geolocation.getCurrentPosition(
          pos => {
            const gps = { lat: pos.coords.latitude, lng: pos.coords.longitude };
            setUserGps(gps);

            if (leafletMapRef.current) {
              if (leafletMarkerRef.current) {
                leafletMapRef.current.removeLayer(leafletMarkerRef.current);
              }
              leafletMarkerRef.current = L.marker([gps.lat, gps.lng])
                .addTo(leafletMapRef.current)
                .bindPopup('Posisi Anda')
                .openPopup();
              leafletMapRef.current.setView([gps.lat, gps.lng], 16);
            }
          },
          err => {
            console.warn('GPS Error:', err);
            setGpsStatusText('GPS Dinonaktifkan');
            setGpsBadgeClass('text-[10px] font-bold px-2 py-0.5 rounded-full bg-rose-100 text-rose-700');
          },
          { enableHighAccuracy: true, timeout: 10000 }
        );
      }, 250);
    }

    if (modalPresensiOpen && presensiStep === 3) {
      startLiveCamera();
    } else {
      stopLiveCamera();
    }
  }, [modalPresensiOpen, presensiStep]);

  // Geofence circle update
  useEffect(() => {
    if (!leafletMapRef.current || cachedLokasi.length === 0) return;
    const targetLoc = cachedLokasi.find(l => l.id === selectedOfficeId) || cachedLokasi[0];

    if (leafletCircleRef.current) {
      leafletMapRef.current.removeLayer(leafletCircleRef.current);
      leafletCircleRef.current = null;
    }

    if (targetLoc) {
      leafletCircleRef.current = L.circle([targetLoc.latitude, targetLoc.longitude], {
        radius: targetLoc.radius_meter,
        color: '#4f46e5',
        fillColor: '#6366f1',
        fillOpacity: 0.2
      }).addTo(leafletMapRef.current);
    }

    // Evaluate distance
    if (!userGps) {
      setGpsStatusText('Menunggu Sinyal GPS');
      setGpsBadgeClass('text-[10px] font-bold px-2 py-0.5 rounded-full bg-amber-100 text-amber-800');
      setIsInsideGeofence(false);
      return;
    }

    if (targetLoc) {
      const dist = calculateDistanceMeters(userGps.lat, userGps.lng, targetLoc.latitude, targetLoc.longitude);
      const isInside = dist <= targetLoc.radius_meter;
      setIsInsideGeofence(isInside);
      if (isInside) {
        setGpsStatusText(`Dalam Radius (${dist}m dari kantor)`);
        setGpsBadgeClass('text-[10px] font-bold px-2 py-0.5 rounded-full bg-emerald-100 text-emerald-700');
      } else {
        setGpsStatusText(`Di Luar Radius (${dist}m / Max ${targetLoc.radius_meter}m)`);
        setGpsBadgeClass('text-[10px] font-bold px-2 py-0.5 rounded-full bg-rose-100 text-rose-700');
      }
    }
  }, [cachedLokasi, selectedOfficeId, userGps]);

  const closeModalPresensi = () => {
    stopLiveCamera();
    setModalPresensiOpen(false);
    if (leafletMapRef.current) {
      leafletMapRef.current.remove();
      leafletMapRef.current = null;
    }
  };

  const submitPresensi = async () => {
    if (!currentUser || !capturedPhoto) return;
    setIsPresensiSubmitting(true);
    try {
      const today = getLocalISODate();
      const officeObj = cachedLokasi.find(l => l.id === selectedOfficeId);
      const photoFileName = `Presensi_${currentUser.username}_${Date.now()}.jpg`;
      const driveUrl = await uploadPhotoToGoogleDrive(capturedPhoto, photoFileName);

      let statusPresensi = 'Hadir';
      let ket: string | null = null;

      if (presensiMode === 'masuk') {
        statusPresensi = 'Hadir';
        const { data: shiftData } = await supabaseClient
          .from('shift_presensi')
          .select('jam_masuk')
          .eq('nama_shift', selectedShiftName)
          .maybeSingle();

        const { data: tolData } = await supabaseClient
          .from('jam_kerja_presensi')
          .select('toleransi_menit')
          .limit(1)
          .maybeSingle();
        const tol = tolData?.toleransi_menit ?? 15;

        if (shiftData?.jam_masuk) {
          const [sh, sm] = shiftData.jam_masuk.split(':').map(Number);
          const now = new Date();
          const nowMenit = now.getHours() * 60 + now.getMinutes();
          const limitMenit = sh * 60 + sm + tol;
          if (nowMenit > limitMenit) {
            statusPresensi = 'Terlambat';
          }
        }
      } else {
        statusPresensi = statusKeluar;
        ket = presensiKeterangan.trim() || null;
      }

      const { error } = await supabaseClient.from('presensi_presensi').insert([{
        karyawan_id: currentUser.id,
        karyawan_nama: currentUser.nama_lengkap,
        divisi: currentUser.divisi,
        tipe: presensiMode,
        status: statusPresensi,
        shift_nama: selectedShiftName,
        lokasi_id: selectedOfficeId || null,
        lokasi_nama: officeObj?.nama_lokasi || null,
        latitude: userGps?.lat || null,
        longitude: userGps?.lng || null,
        foto_url: driveUrl,
        keterangan: ket,
        tanggal: today
      }]);

      if (error) throw error;

      showToast(`Presensi ${presensiMode} (${statusPresensi}) berhasil disimpan!`, 'success');
      closeModalPresensi();
      refreshUserDashboard();
      loadEmployeeDailyStatus();
    } catch (err) {
      console.error(err);
      showToast('Terjadi kesalahan saat menyimpan presensi.', 'error');
    } finally {
      setIsPresensiSubmitting(false);
    }
  };

  // Submit Lembut Action
  const submitLembur = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!currentUser) return;
    const durasi = parseInt(lemburDurasi, 10);
    if (!durasi || durasi < 1) {
      showToast('Silakan masukkan durasi lembur (minimal 1 jam).', 'error');
      return;
    }
    setLemburSubmitting(true);
    try {
      let buktiUrl: string | null = null;
      if (lemburFoto) {
        buktiUrl = await uploadPhotoToGoogleDrive(lemburFoto, `Lembur_${currentUser.username}_${Date.now()}.jpg`);
      }

      const { error } = await supabaseClient.from('lembur_presensi').insert([{
        karyawan_id: currentUser.id,
        karyawan_nama: currentUser.nama_lengkap,
        durasi_jam: durasi,
        keterangan: lemburKeterangan.trim(),
        bukti_url: buktiUrl,
        status: 'Pending',
        tanggal: getLocalISODate()
      }]);

      if (error) throw error;

      showToast('Pengajuan lembur berhasil dikirim!', 'success');
      setModalLemburOpen(false);
      setLemburFoto(null);
      setLemburDurasi('');
      setLemburKeterangan('');
    } catch {
      showToast('Gagal mengajukan lembur.', 'error');
    } finally {
      setLemburSubmitting(false);
    }
  };

  // Submit Izin / Sakit / Cuti
  const submitIzin = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!currentUser) return;
    setIzinSubmitting(true);
    try {
      let driveUrl: string | null = null;
      if (izinFoto) {
        driveUrl = await uploadPhotoToGoogleDrive(izinFoto, `${izinType}_${currentUser.username}_${Date.now()}.jpg`);
      }

      const today = getLocalISODate();
      const userShift = (currentUser.shifts && currentUser.shifts[0]) || 'Reguler';
      const { error } = await supabaseClient.from('presensi_presensi').insert([{
        karyawan_id: currentUser.id,
        karyawan_nama: currentUser.nama_lengkap,
        divisi: currentUser.divisi,
        tipe: 'keluar',
        status: izinType,
        shift_nama: userShift,
        foto_url: driveUrl,
        keterangan: izinKeterangan.trim(),
        tanggal: today
      }]);

      if (error) throw error;

      if (izinType === 'CUTI') {
        const curCuti = currentUser.sisa_cuti ?? 12;
        const newCuti = Math.max(0, curCuti - 1);
        await supabaseClient.from('karyawan_presensi').update({ sisa_cuti: newCuti }).eq('id', currentUser.id);
        setCurrentUser(prev => prev ? { ...prev, sisa_cuti: newCuti } : null);
      }

      const label = izinType === 'SAKIT' ? 'Sakit' : (izinType === 'IZIN' ? 'Izin' : 'Cuti');
      showToast(`Pengajuan ${label} berhasil disimpan!`, 'success');
      setModalIzinOpen(false);
      setIzinFoto(null);
      setIzinKeterangan('');
      refreshUserDashboard();
      loadEmployeeDailyStatus();
    } catch {
      showToast(`Gagal menyimpan pengajuan ${izinType}.`, 'error');
    } finally {
      setIzinSubmitting(false);
    }
  };

  // Submit Libur Direct Action
  const submitLiburDirect = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!currentUser) return;
    setLiburSubmitting(true);
    try {
      const today = getLocalISODate();
      const userShift = (currentUser.shifts && currentUser.shifts[0]) || 'Reguler';
      const { error } = await supabaseClient.from('presensi_presensi').insert([{
        karyawan_id: currentUser.id,
        karyawan_nama: currentUser.nama_lengkap,
        divisi: currentUser.divisi,
        tipe: 'keluar',
        status: 'LIBUR',
        shift_nama: userShift,
        keterangan: liburKeterangan.trim() || 'Jadwal Libur',
        tanggal: today
      }]);

      if (error) throw error;

      const curLibur = currentUser.sisa_libur ?? 4;
      const newLibur = Math.max(0, curLibur - 1);
      await supabaseClient.from('karyawan_presensi').update({ sisa_libur: newLibur }).eq('id', currentUser.id);
      setCurrentUser(prev => prev ? { ...prev, sisa_libur: newLibur } : null);

      showToast('Presensi Libur berhasil dicatat!', 'success');
      setModalLiburOpen(false);
      setLiburKeterangan('');
      refreshUserDashboard();
      loadEmployeeDailyStatus();
    } catch {
      showToast('Gagal mencatat presensi libur.', 'error');
    } finally {
      setLiburSubmitting(false);
    }
  };

  // Handle Cuti button click with allowance check
  const handleMenuCutiClick = async () => {
    if (!currentUser) return;
    if (currentUser.divisi) {
      try {
        const { data: div } = await supabaseClient
          .from('divisi_presensi')
          .select('allow_cuti')
          .eq('nama_divisi', currentUser.divisi)
          .maybeSingle();

        if (div && div.allow_cuti === false) {
          setModalCutiLockedOpen(true);
          return;
        }
      } catch {}
    }
    setIzinType('CUTI');
    setIzinFoto(null);
    setIzinKeterangan('');
    setModalIzinOpen(true);
  };

  // Photo viewer modal
  const openPhotoPreviewModal = (url: string, rawNama: string, rawSubtitle: string) => {
    setPreviewPhotoUrl(formatDirectImageUrl(url));
    setPreviewPhotoTitle(decodeURIComponent(rawNama || 'Foto Presensi'));
    setPreviewPhotoSubtitle(decodeURIComponent(rawSubtitle || ''));
    setPreviewPhotoLoading(true);
    setPreviewPhotoError(false);
    setPhotoPreviewModalOpen(true);
  };

  // Date shifter navigation
  const shiftRekapTanggal = (days: number) => {
    const base = new Date(filterRekapTanggal || getLocalISODate());
    base.setDate(base.getDate() + days);
    const pad = (n: number) => String(n).padStart(2, '0');
    const nextIso = `${base.getFullYear()}-${pad(base.getMonth() + 1)}-${pad(base.getDate())}`;
    setFilterRekapTanggal(nextIso);
    loadAdminRekapan(nextIso);
  };

  const openExportExcelModal = () => {
    setExportStartDate(filterRekapTanggal || getLocalISODate());
    setExportEndDate(filterRekapTanggal || getLocalISODate());
    setModalExportExcelOpen(true);
  };

  const handleDownloadExcelPeriod = async () => {
    if (!exportStartDate || !exportEndDate) {
      showToast('Harap tentukan tanggal mulai dan tanggal selesai.', 'error');
      return;
    }
    if (exportStartDate > exportEndDate) {
      showToast('Tanggal mulai tidak boleh melebihi tanggal selesai.', 'error');
      return;
    }

    setExportExcelLoading(true);
    try {
      const { data: recs, error } = await supabaseClient
        .from('presensi_presensi')
        .select('*')
        .gte('tanggal', exportStartDate)
        .lte('tanggal', exportEndDate)
        .order('tanggal', { ascending: true })
        .order('created_at', { ascending: true });

      if (error) throw error;

      const records = recs || [];
      if (records.length === 0) {
        showToast('Tidak ada data presensi pada rentang tanggal yang dipilih.', 'warning');
        setExportExcelLoading(false);
        return;
      }

      // Gabungkan presensi masuk dan keluar per karyawan per tanggal
      const mapDayEmp = new Map<string, {
        tanggal: string;
        karyawan_nama: string;
        divisi: string;
        shift_nama: string;
        jam_masuk: string;
        jam_keluar: string;
        status: string;
        lokasi: string;
        keterangan: string;
        foto_masuk: string;
        foto_keluar: string;
      }>();

      records.forEach(r => {
        const key = `${r.tanggal}_${r.karyawan_id || r.karyawan_nama}`;
        if (!mapDayEmp.has(key)) {
          mapDayEmp.set(key, {
            tanggal: r.tanggal,
            karyawan_nama: r.karyawan_nama,
            divisi: r.divisi || '-',
            shift_nama: r.shift_nama || '-',
            jam_masuk: '-',
            jam_keluar: '-',
            status: r.status || 'Hadir',
            lokasi: r.lokasi_nama || '-',
            keterangan: r.keterangan || '-',
            foto_masuk: '-',
            foto_keluar: '-'
          });
        }
        const item = mapDayEmp.get(key)!;
        if (r.tipe === 'masuk') {
          item.jam_masuk = formatTime24(r.created_at);
          if (r.foto_url) item.foto_masuk = r.foto_url;
          if (r.status === 'Terlambat') item.status = 'Terlambat';
        } else {
          item.jam_keluar = formatTime24(r.created_at);
          if (r.foto_url) item.foto_keluar = r.foto_url;
          item.status = r.status || item.status;
        }
        if (r.keterangan && item.keterangan === '-') item.keterangan = r.keterangan;
        else if (r.keterangan && item.keterangan !== r.keterangan) item.keterangan += ` | ${r.keterangan}`;
        if (r.lokasi_nama && item.lokasi === '-') item.lokasi = r.lokasi_nama;
      });

      const excelRows = Array.from(mapDayEmp.values()).map((row, idx) => [
        idx + 1,
        row.tanggal,
        row.karyawan_nama,
        row.divisi,
        row.shift_nama,
        row.jam_masuk,
        row.jam_keluar,
        row.status,
        row.lokasi,
        row.keterangan,
        row.foto_masuk,
        row.foto_keluar
      ]);

      const headers = [
        'No',
        'Tanggal',
        'Nama Karyawan',
        'Divisi',
        'Shift',
        'Jam Masuk',
        'Jam Keluar',
        'Status',
        'Lokasi',
        'Keterangan',
        'Link Foto Masuk',
        'Link Foto Keluar'
      ];

      const ws = XLSX.utils.aoa_to_sheet([headers, ...excelRows]);
      ws['!cols'] = [
        { wch: 6 },
        { wch: 14 },
        { wch: 25 },
        { wch: 18 },
        { wch: 15 },
        { wch: 12 },
        { wch: 12 },
        { wch: 16 },
        { wch: 20 },
        { wch: 30 },
        { wch: 40 },
        { wch: 40 }
      ];

      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Rekap Presensi');
      XLSX.writeFile(wb, `Rekap_Presensi_${exportStartDate}_sd_${exportEndDate}.xlsx`);

      showToast(`Berhasil mengunduh Excel periode ${exportStartDate} s/d ${exportEndDate}!`, 'success');
      setModalExportExcelOpen(false);
    } catch (err: any) {
      console.error(err);
      showToast('Gagal mengunduh file Excel: ' + (err.message || 'Terjadi kesalahan'), 'error');
    } finally {
      setExportExcelLoading(false);
    }
  };

  // Login Handlers
  const [karyawanLoginId, setKaryawanLoginId] = useState('');
  const [karyawanLoginPass, setKaryawanLoginPass] = useState('');
  const [karyawanLoginLoading, setKaryawanLoginLoading] = useState(false);

  const handleKaryawanLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    const id = karyawanLoginId.trim();
    const pass = karyawanLoginPass.trim();
    setKaryawanLoginLoading(true);

    try {
      const { data: users, error } = await supabaseClient
        .from('karyawan_presensi')
        .select('*')
        .or(`username.eq.${id},no_wa.eq.${id},email.eq.${id}`);

      if (error) throw error;

      const match = users?.find(u => u.password === pass);
      if (!match) {
        showToast('ID Karyawan atau password salah!', 'error');
        return;
      }

      setCurrentUser(match);
      localStorage.setItem('presensi_user', JSON.stringify(match));
      showToast(`Selamat datang, ${match.nama_lengkap}!`, 'success');
      setKaryawanLoginId('');
      setKaryawanLoginPass('');
    } catch {
      showToast('Gagal terhubung ke database!', 'error');
    } finally {
      setKaryawanLoginLoading(false);
    }
  };

  const logoutKaryawan = () => {
    setCurrentUser(null);
    localStorage.removeItem('presensi_user');
    navigateTo('#karyawan');
    showToast('Berhasil keluar dari mode karyawan.', 'info');
  };

  const [adminLoginId, setAdminLoginId] = useState('');
  const [adminLoginPass, setAdminLoginPass] = useState('');
  const [adminLoginLoading, setAdminLoginLoading] = useState(false);

  const handleAdminLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    const id = adminLoginId.trim();
    const pass = adminLoginPass.trim();
    setAdminLoginLoading(true);

    try {
      const { data: admins, error } = await supabaseClient
        .from('admin_presensi')
        .select('*')
        .or(`username.ilike.${id},email.ilike.${id}`);

      if (error) throw error;

      const match = admins?.find(a => a.password === pass);
      if (!match) {
        showToast('Kredensial Admin tidak cocok!', 'error');
        return;
      }

      setCurrentAdmin(match);
      localStorage.setItem('presensi_admin', JSON.stringify(match));
      showToast('Login Admin Berhasil!', 'success');
      setAdminLoginId('');
      setAdminLoginPass('');
      switchAdminTab('rekap');
    } catch {
      showToast('Gagal memverifikasi akun admin.', 'error');
    } finally {
      setAdminLoginLoading(false);
    }
  };

  const logoutAdmin = () => {
    setCurrentAdmin(null);
    localStorage.removeItem('presensi_admin');
    showToast('Admin berhasil keluar.', 'info');
  };

  // Grouped history dates for employee view
  const groupedUserHistory = React.useMemo(() => {
    const mapByDate = new Map<string, PresensiRecord[]>();
    userHistoryRecords.forEach(r => {
      const tgl = r.tanggal;
      if (!tgl) return;
      if (!mapByDate.has(tgl)) mapByDate.set(tgl, []);
      mapByDate.get(tgl)!.push(r);
    });
    return Array.from(mapByDate.entries()).sort((a, b) => b[0].localeCompare(a[0]));
  }, [userHistoryRecords]);

  // Attendance buttons states for today
  const hasLiburToday = todayAttendanceRecords.some(r => (r.status || '').toUpperCase() === 'LIBUR');
  const hasIzinToday = todayAttendanceRecords.some(r => (r.status || '').toUpperCase() === 'IZIN');
  const masukRecordToday = todayAttendanceRecords.find(r => r.tipe === 'masuk');
  const keluarRecordToday = todayAttendanceRecords.find(r => r.tipe === 'keluar');

  // Filter counters for Team daily view
  const teamCounters = React.useMemo(() => {
    const c = { masuk: 0, belum_masuk: 0, izin: 0, libur: 0, cuti: 0 };
    teamDailyData.forEach(item => {
      if (c[item.category] !== undefined) c[item.category]++;
    });
    return c;
  }, [teamDailyData]);

  return (
    <div className="bg-slate-100 text-slate-800 antialiased min-h-screen flex flex-col font-sans">
      {/* Header Bar */}
      <header className="bg-white/80 backdrop-blur-md sticky top-0 z-40 border-b border-slate-200 shadow-sm transition-all">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 py-2.5 flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            <div className="w-9 h-9 rounded-xl bg-gradient-to-tr from-brand-600 to-indigo-500 text-white flex items-center justify-center shadow-md shadow-brand-500/20">
              <i className="fa-solid fa-camera text-base"></i>
            </div>
            <div>
              <h1 className="text-lg font-extrabold tracking-tight text-slate-900 leading-none">Presensi</h1>
              <span className="text-[11px] font-semibold text-brand-600 tracking-wide uppercase">
                {currentView === 'dashboard' ? 'Dashboard Admin' : (currentView === 'status-karyawan' ? 'Status Karyawan' : 'Mode Karyawan')}
              </span>
            </div>
          </div>

          {/* Switcher Bar */}
          <div className="flex items-center gap-1.5 bg-slate-100 p-1 rounded-xl border border-slate-200 text-xs font-semibold">
            <button
              onClick={() => navigateTo('#karyawan')}
              className={`px-3 py-1.5 rounded-lg transition-all flex items-center gap-1.5 cursor-pointer ${
                currentView === 'karyawan' ? 'bg-white text-brand-600 shadow-sm' : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <i className="fa-solid fa-mobile-screen"></i> Karyawan
            </button>
            <button
              onClick={() => navigateTo('#dashboard')}
              className={`px-3 py-1.5 rounded-lg transition-all flex items-center gap-1.5 cursor-pointer ${
                currentView === 'dashboard' ? 'bg-white text-brand-600 shadow-sm' : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <i className="fa-solid fa-gauge-high"></i> Dashboard
            </button>
          </div>
        </div>
      </header>

      {/* Main Container */}
      <main className="flex-1 w-full relative flex flex-col">
        {/* VIEW 1: MODE KARYAWAN */}
        {currentView === 'karyawan' && (
          <section className="flex-1 max-w-md w-full mx-auto p-4 flex flex-col gap-4">
            {!currentUser ? (
              /* Karyawan Login Form */
              <div className="bg-white p-6 rounded-3xl shadow-xl shadow-slate-200/60 border border-slate-100 my-auto">
                <div className="text-center mb-6">
                  <div className="w-16 h-16 rounded-2xl bg-brand-50 text-brand-600 flex items-center justify-center mx-auto mb-3 shadow-inner">
                    <i className="fa-solid fa-camera text-2xl"></i>
                  </div>
                  <h2 className="text-2xl font-extrabold text-slate-900">Selamat Datang!</h2>
                  <p className="text-xs text-slate-500 mt-1">Silakan masuk menggunakan akun karyawan Anda</p>
                </div>

                <form onSubmit={handleKaryawanLogin} className="space-y-4">
                  <div>
                    <label className="block text-xs font-bold text-slate-600 mb-1.5 uppercase tracking-wider">ID Karyawan</label>
                    <div className="relative">
                      <span className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none text-slate-400">
                        <i className="fa-solid fa-user-tag text-sm"></i>
                      </span>
                      <input
                        type="text"
                        value={karyawanLoginId}
                        onChange={e => setKaryawanLoginId(e.target.value)}
                        required
                        placeholder="Username / No. whatsapp / email"
                        className="w-full pl-10 pr-4 py-3 bg-slate-50 border border-slate-200 rounded-2xl text-sm focus:outline-none focus:ring-2 focus:ring-brand-500/20 focus:border-brand-500 font-medium transition-all placeholder:text-slate-400 placeholder:text-xs"
                      />
                    </div>
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-600 mb-1.5 uppercase tracking-wider">Password</label>
                    <div className="relative">
                      <span className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none text-slate-400">
                        <i className="fa-solid fa-lock text-sm"></i>
                      </span>
                      <input
                        type="password"
                        value={karyawanLoginPass}
                        onChange={e => setKaryawanLoginPass(e.target.value)}
                        required
                        placeholder="Masukkan password"
                        className="w-full pl-10 pr-4 py-3 bg-slate-50 border border-slate-200 rounded-2xl text-sm focus:outline-none focus:ring-2 focus:ring-brand-500/20 focus:border-brand-500 font-medium transition-all"
                      />
                    </div>
                  </div>

                  <button
                    type="submit"
                    disabled={karyawanLoginLoading}
                    className="w-full py-3.5 px-4 bg-brand-600 hover:bg-brand-700 active:scale-[0.99] text-white font-bold rounded-2xl shadow-lg shadow-brand-500/30 transition-all flex items-center justify-center gap-2 cursor-pointer disabled:opacity-60"
                  >
                    {karyawanLoginLoading ? (
                      <>
                        <i className="fa-solid fa-spinner fa-spin"></i>
                        <span>Memeriksa...</span>
                      </>
                    ) : (
                      <span>Masuk</span>
                    )}
                  </button>
                </form>
              </div>
            ) : (
              /* Karyawan Dashboard Box */
              <div className="flex flex-col gap-4 pb-12">
                {/* Kartu Profil */}
                <div className="bg-gradient-to-br from-brand-600 via-indigo-600 to-brand-700 text-white rounded-3xl p-5 shadow-xl shadow-brand-500/25 relative overflow-hidden">
                  <div className="absolute -right-6 -bottom-6 w-32 h-32 rounded-full bg-white/10 blur-xl pointer-events-none"></div>

                  <div className="flex items-center justify-between mb-4">
                    <div className="flex items-center gap-3">
                      <div className="w-12 h-12 rounded-2xl bg-white/20 backdrop-blur-md flex items-center justify-center border border-white/30 text-xl font-black">
                        <span>{currentUser.nama_lengkap.charAt(0).toUpperCase()}</span>
                      </div>
                      <div>
                        <h3 className="font-extrabold text-base leading-tight">{currentUser.nama_lengkap}</h3>
                        <span className="text-xs text-indigo-100 font-semibold">{currentUser.divisi || 'Umum'}</span>
                      </div>
                    </div>
                    <div className="flex items-center gap-1.5">
                      <button
                        type="button"
                        onClick={() => setModalAllPengumumanOpen(true)}
                        className="relative w-9 h-9 rounded-xl bg-white/15 hover:bg-white/25 active:scale-95 flex items-center justify-center text-white transition-all text-xs cursor-pointer"
                        title="Papan Pengumuman Kantor"
                      >
                        <i className="fa-solid fa-bullhorn"></i>
                        {karyawanPengumumanList.length > 0 && (
                          <span className="absolute -top-1 -right-1 w-4 h-4 bg-amber-400 text-slate-900 rounded-full text-[9px] font-black flex items-center justify-center border-2 border-brand-700 animate-pulse">
                            {karyawanPengumumanList.length}
                          </span>
                        )}
                      </button>
                      <button
                        type="button"
                        onClick={triggerAddToHomeScreen}
                        className="w-9 h-9 rounded-xl bg-white/15 hover:bg-white/25 active:scale-95 flex items-center justify-center text-white transition-all text-xs cursor-pointer"
                        title="Pasang ke Layar Utama (Fullscreen)"
                      >
                        <i className="fa-solid fa-mobile-screen-button"></i>
                      </button>
                      <button
                        onClick={logoutKaryawan}
                        className="w-9 h-9 rounded-xl bg-white/15 hover:bg-white/25 active:scale-95 flex items-center justify-center text-white transition-all text-xs cursor-pointer"
                        title="Keluar"
                      >
                        <i className="fa-solid fa-arrow-right-from-bracket"></i>
                      </button>
                    </div>
                  </div>

                  {/* Jam Digital Real-time & Tanggal */}
                  <div className="bg-black/20 rounded-2xl p-3.5 backdrop-blur-sm border border-white/10 flex flex-col items-center justify-center text-center">
                    <div className="text-3xl font-extrabold tracking-wider font-mono">{currentTime || '00:00:00'}</div>
                    <div className="text-xs text-indigo-100 font-medium mt-0.5">{currentDateStr || 'Hari ini'}</div>

                    {/* Badge Info Jam Kerja Shift */}
                    <div className={todayShiftInfo.className}>
                      <i className={todayShiftInfo.icon}></i>
                      <span>{todayShiftInfo.text}</span>
                    </div>
                  </div>

                  {/* Jatah Libur & Cuti Badge */}
                  <div className="grid grid-cols-2 gap-2 mt-3 pt-3 border-t border-white/10">
                    <div className="flex items-center gap-2">
                      <div className="w-7 h-7 rounded-lg bg-white/10 flex items-center justify-center text-xs">
                        <i className="fa-solid fa-mug-hot"></i>
                      </div>
                      <div>
                        <div className="text-[10px] text-indigo-200 uppercase font-semibold">Sisa Libur</div>
                        <div className="text-sm font-extrabold">{currentUser.sisa_libur ?? 4} Hari</div>
                      </div>
                    </div>
                    <div className="flex items-center gap-2">
                      <div className="w-7 h-7 rounded-lg bg-white/10 flex items-center justify-center text-xs">
                        <i className="fa-solid fa-plane-departure"></i>
                      </div>
                      <div>
                        <div className="text-[10px] text-indigo-200 uppercase font-semibold">Sisa Cuti</div>
                        <div className="text-sm font-extrabold">{currentUser.sisa_cuti ?? 12} Hari</div>
                      </div>
                    </div>
                  </div>
                </div>

                {/* Banner Pengumuman Kantor */}
                {karyawanPengumumanList.length > 0 && (
                  <div className="bg-gradient-to-br from-amber-50 via-orange-50/40 to-amber-50/60 border border-amber-200/90 rounded-3xl p-4 shadow-sm space-y-2.5">
                    <div className="flex items-center justify-between">
                      <div className="flex items-center gap-2">
                        <div className="w-7 h-7 rounded-xl bg-amber-500 text-white flex items-center justify-center text-xs shadow-sm">
                          <i className="fa-solid fa-bullhorn"></i>
                        </div>
                        <span className="text-xs font-black text-amber-950 uppercase tracking-wider">
                          Pengumuman Kantor
                        </span>
                        <span className="px-2 py-0.5 rounded-full text-[10px] font-extrabold bg-amber-200 text-amber-800">
                          Terbaru
                        </span>
                      </div>
                      {karyawanPengumumanList.length > 1 && (
                        <button
                          type="button"
                          onClick={() => setModalAllPengumumanOpen(true)}
                          className="text-[11px] font-bold text-amber-800 hover:text-amber-950 flex items-center gap-1 cursor-pointer"
                        >
                          <span>Semua ({karyawanPengumumanList.length})</span>
                          <i className="fa-solid fa-chevron-right text-[9px]"></i>
                        </button>
                      )}
                    </div>

                    <div
                      onClick={() => setModalPengumumanDetail(karyawanPengumumanList[0])}
                      className="p-3 bg-white/90 backdrop-blur-sm rounded-2xl border border-amber-200/70 hover:bg-white hover:border-amber-300 transition-all cursor-pointer group shadow-2xs"
                    >
                      <div className="flex items-start justify-between gap-2">
                        <h4 className="font-extrabold text-xs text-slate-900 group-hover:text-brand-600 transition-colors line-clamp-1">
                          {karyawanPengumumanList[0].judul}
                        </h4>
                        <span className="text-[10px] text-slate-400 font-semibold shrink-0">
                          {formatIndoDate(karyawanPengumumanList[0].created_at?.split('T')[0] || '')}
                        </span>
                      </div>
                      <p className="text-xs text-slate-600 mt-1 line-clamp-2 leading-relaxed font-normal">
                        {karyawanPengumumanList[0].isi}
                      </p>
                      <div className="mt-2.5 flex items-center justify-between text-[10px] text-slate-500 font-medium pt-1.5 border-t border-slate-100">
                        <span className="flex items-center gap-1 text-slate-600 font-semibold">
                          <i className="fa-solid fa-user-shield text-[10px] text-amber-600"></i>
                          {karyawanPengumumanList[0].author || 'Admin Kantor'}
                        </span>
                        <span className="text-brand-600 font-bold flex items-center gap-1 group-hover:translate-x-0.5 transition-transform">
                          <span>Baca Selengkapnya</span>
                          <i className="fa-solid fa-arrow-right text-[9px]"></i>
                        </span>
                      </div>
                    </div>
                  </div>
                )}

                {/* Tombol Presensi Utama (Masuk & Keluar) */}
                <div className="grid grid-cols-2 gap-3">
                  {/* Absen Masuk */}
                  {hasLiburToday || hasIzinToday ? (
                    <button
                      disabled
                      className="p-4 bg-slate-200 text-slate-400 rounded-3xl border border-slate-300 flex flex-col items-center justify-center gap-2 cursor-not-allowed pointer-events-none opacity-70 select-none"
                    >
                      <div className="w-12 h-12 rounded-2xl bg-slate-300 text-slate-500 flex items-center justify-center text-xl">
                        <i className="fa-solid fa-arrow-right-to-bracket"></i>
                      </div>
                      <span className="font-extrabold text-sm text-slate-500">Absen Masuk</span>
                      <span className="text-[11px] text-slate-400 font-medium">{hasLiburToday ? 'Libur' : 'Izin'}</span>
                    </button>
                  ) : masukRecordToday ? (
                    <button
                      disabled
                      className="p-4 bg-slate-200 text-slate-400 rounded-3xl border border-slate-300 flex flex-col items-center justify-center gap-2 cursor-not-allowed pointer-events-none opacity-80 select-none"
                    >
                      <div className="w-12 h-12 rounded-2xl bg-slate-300 text-slate-500 flex items-center justify-center text-xl">
                        <i className="fa-solid fa-circle-check text-emerald-600"></i>
                      </div>
                      <span className="font-extrabold text-sm text-slate-600">Sudah Masuk</span>
                      <span className="text-[11px] font-mono font-bold text-slate-500">
                        Pukul {formatTime24(masukRecordToday.created_at)} WIB
                      </span>
                    </button>
                  ) : (
                    <button
                      onClick={() => openModalPresensi('masuk')}
                      className="p-4 bg-emerald-600 hover:bg-emerald-700 active:scale-[0.98] text-white rounded-3xl shadow-lg shadow-emerald-500/25 flex flex-col items-center justify-center gap-2 transition-all cursor-pointer"
                    >
                      <div className="w-12 h-12 rounded-2xl bg-white/20 flex items-center justify-center text-xl">
                        <i className="fa-solid fa-arrow-right-to-bracket"></i>
                      </div>
                      <span className="font-extrabold text-sm">Absen Masuk</span>
                    </button>
                  )}

                  {/* Absen Keluar */}
                  {hasLiburToday || hasIzinToday ? (
                    <button
                      disabled
                      className="p-4 bg-slate-200 text-slate-400 rounded-3xl border border-slate-300 flex flex-col items-center justify-center gap-2 cursor-not-allowed pointer-events-none opacity-70 select-none"
                    >
                      <div className="w-12 h-12 rounded-2xl bg-slate-300 text-slate-500 flex items-center justify-center text-xl">
                        <i className="fa-solid fa-arrow-right-from-bracket"></i>
                      </div>
                      <span className="font-extrabold text-sm text-slate-500">Absen Keluar</span>
                      <span className="text-[11px] text-slate-400 font-medium">{hasLiburToday ? 'Libur' : 'Izin'}</span>
                    </button>
                  ) : keluarRecordToday ? (
                    <button
                      disabled
                      className="p-4 bg-slate-200 text-slate-400 rounded-3xl border border-slate-300 flex flex-col items-center justify-center gap-2 cursor-not-allowed pointer-events-none opacity-80 select-none"
                    >
                      <div className="w-12 h-12 rounded-2xl bg-slate-300 text-slate-500 flex items-center justify-center text-xl">
                        <i className="fa-solid fa-circle-check text-rose-600"></i>
                      </div>
                      <span className="font-extrabold text-sm text-slate-600">Sudah Keluar</span>
                      <span className="text-[11px] font-mono font-bold text-slate-500">
                        {keluarRecordToday.status} • {formatTime24(keluarRecordToday.created_at)}
                      </span>
                    </button>
                  ) : (
                    <button
                      onClick={() => openModalPresensi('keluar')}
                      className="p-4 bg-rose-600 hover:bg-rose-700 active:scale-[0.98] text-white rounded-3xl shadow-lg shadow-rose-500/25 flex flex-col items-center justify-center gap-2 transition-all cursor-pointer"
                    >
                      <div className="w-12 h-12 rounded-2xl bg-white/20 flex items-center justify-center text-xl">
                        <i className="fa-solid fa-arrow-right-from-bracket"></i>
                      </div>
                      <span className="font-extrabold text-sm">Absen Keluar</span>
                    </button>
                  )}
                </div>

                {/* Tombol Aksi Sekunder: Sakit, Libur/Izin, Cuti, Lembur */}
                <div className="grid grid-cols-4 gap-2">
                  <button
                    onClick={() => {
                      setIzinType('SAKIT');
                      setIzinFoto(null);
                      setIzinKeterangan('');
                      setModalIzinOpen(true);
                    }}
                    className="p-3 bg-white hover:bg-slate-50 border border-slate-200 rounded-2xl flex flex-col items-center gap-1.5 text-slate-700 active:scale-95 transition-all shadow-sm cursor-pointer"
                  >
                    <div className="w-8 h-8 rounded-xl bg-amber-50 text-amber-600 flex items-center justify-center text-sm">
                      <i className="fa-solid fa-briefcase-medical"></i>
                    </div>
                    <span className="text-[11px] font-bold">Sakit</span>
                  </button>

                  <button
                    onClick={() => {
                      if ((currentUser.sisa_libur ?? 0) <= 0) {
                        setIzinType('IZIN');
                        setIzinFoto(null);
                        setIzinKeterangan('');
                        setModalIzinOpen(true);
                      } else {
                        setLiburKeterangan('');
                        setModalLiburOpen(true);
                      }
                    }}
                    className="p-3 bg-white hover:bg-slate-50 border border-slate-200 rounded-2xl flex flex-col items-center gap-1.5 text-slate-700 active:scale-95 transition-all shadow-sm cursor-pointer"
                  >
                    <div className={`w-8 h-8 rounded-xl flex items-center justify-center text-sm ${
                      (currentUser.sisa_libur ?? 0) <= 0 ? 'bg-amber-50 text-amber-600' : 'bg-sky-50 text-sky-600'
                    }`}>
                      <i className={`fa-solid ${(currentUser.sisa_libur ?? 0) <= 0 ? 'fa-envelope-open-text' : 'fa-umbrella-beach'}`}></i>
                    </div>
                    <span className="text-[11px] font-bold">
                      {(currentUser.sisa_libur ?? 0) <= 0 ? 'Izin' : 'Libur'}
                    </span>
                  </button>

                  <button
                    onClick={handleMenuCutiClick}
                    className="p-3 bg-white hover:bg-slate-50 border border-slate-200 rounded-2xl flex flex-col items-center gap-1.5 text-slate-700 active:scale-95 transition-all shadow-sm cursor-pointer"
                  >
                    <div className="w-8 h-8 rounded-xl bg-purple-50 text-purple-600 flex items-center justify-center text-sm">
                      <i className="fa-solid fa-calendar-check"></i>
                    </div>
                    <span className="text-[11px] font-bold">Cuti</span>
                  </button>

                  <button
                    onClick={() => {
                      setLemburFoto(null);
                      setLemburDurasi('');
                      setLemburKeterangan('');
                      setModalLemburOpen(true);
                    }}
                    className="p-3 bg-white hover:bg-slate-50 border border-slate-200 rounded-2xl flex flex-col items-center gap-1.5 text-slate-700 active:scale-95 transition-all shadow-sm cursor-pointer"
                  >
                    <div className="w-8 h-8 rounded-xl bg-orange-50 text-orange-600 flex items-center justify-center text-sm">
                      <i className="fa-solid fa-clock-rotate-left"></i>
                    </div>
                    <span className="text-[11px] font-bold">Lembur</span>
                  </button>
                </div>

                {/* Tombol Menu: Halaman Status Karyawan Hari Ini */}
                <button
                  type="button"
                  onClick={() => navigateTo('#status-karyawan')}
                  className="w-full p-4 bg-white hover:bg-slate-50 border border-slate-200 rounded-3xl shadow-sm flex items-center justify-between transition-all active:scale-[0.99] group cursor-pointer"
                >
                  <div className="flex items-center gap-3">
                    <div className="w-10 h-10 rounded-2xl bg-brand-50 text-brand-600 flex items-center justify-center text-base group-hover:scale-105 transition-transform shadow-inner">
                      <i className="fa-solid fa-users-viewfinder"></i>
                    </div>
                    <div className="text-left">
                      <div className="font-extrabold text-xs text-slate-900 uppercase tracking-wider">Status Karyawan Hari Ini</div>
                      <div className="text-[11px] text-slate-500 font-medium mt-0.5">Lihat aktivitas & status seluruh rekan kerja</div>
                    </div>
                  </div>
                  <div className="w-8 h-8 rounded-xl bg-slate-100 flex items-center justify-center text-slate-400 group-hover:text-brand-600 group-hover:bg-brand-50 transition-all">
                    <i className="fa-solid fa-chevron-right text-xs"></i>
                  </div>
                </button>

                {/* Riwayat Aktivitas Periode Bulan Aktif (26 - 25) */}
                <div className="bg-white rounded-3xl p-4 border border-slate-200 shadow-sm">
                  <div className="flex items-center justify-between mb-3">
                    <div>
                      <h4 className="font-extrabold text-xs text-slate-900 uppercase tracking-wider">Aktivitas Periode Bulan Aktif</h4>
                      <p className="text-[10px] text-slate-400 font-medium">Siklus tanggal 26 - 25</p>
                    </div>
                    <span className="text-[10px] font-bold text-brand-700 bg-brand-50 px-2.5 py-1 rounded-xl border border-brand-100">
                      {periodLabel}
                    </span>
                  </div>

                  <div className="space-y-2">
                    {groupedUserHistory.length === 0 ? (
                      <div className="text-xs text-slate-400 text-center py-4">Belum ada catatan aktivitas pada periode ini.</div>
                    ) : (
                      groupedUserHistory.map(([tgl, dayRecs]) => {
                        const act = formatDailyActivityText(dayRecs);
                        let dotColor = 'bg-emerald-500';
                        if (act.type === 'libur') dotColor = 'bg-sky-500';
                        else if (act.type === 'cuti') dotColor = 'bg-purple-500';
                        else if (act.type === 'sakit') dotColor = 'bg-amber-500';
                        else if (act.type === 'izin') dotColor = 'bg-orange-500';

                        return (
                          <div
                            key={tgl}
                            className="flex items-start sm:items-center justify-between p-2.5 rounded-2xl bg-slate-50 border border-slate-100 text-xs hover:bg-slate-100/70 transition-all"
                          >
                            <div className="flex items-start sm:items-center gap-2 min-w-0 pr-2">
                              <span className={`w-2 h-2 rounded-full ${dotColor} shrink-0 mt-1 sm:mt-0`}></span>
                              <div className="text-slate-800 text-xs leading-relaxed">
                                <span className="font-bold text-slate-900">{formatIndoDate(tgl)}</span>
                                <span className="text-slate-500 font-bold mx-1">:</span>
                                <span className="font-semibold text-slate-700">{act.text}</span>
                              </div>
                            </div>
                          </div>
                        );
                      })
                    )}
                  </div>
                </div>
              </div>
            )}
          </section>
        )}

        {/* VIEW 2: STATUS KARYAWAN HARI INI */}
        {currentView === 'status-karyawan' && (
          <section className="flex-1 max-w-md w-full mx-auto p-4 flex flex-col gap-4 pb-12">
            {/* Top Navigation */}
            <div className="flex items-center justify-between">
              <button
                type="button"
                onClick={() => navigateTo('#karyawan')}
                className="px-3.5 py-2 bg-white hover:bg-slate-50 border border-slate-200 rounded-2xl text-xs font-bold text-slate-700 flex items-center gap-2 shadow-sm transition-all active:scale-95 cursor-pointer"
              >
                <i className="fa-solid fa-arrow-left text-brand-600"></i>
                <span>Kembali</span>
              </button>
              <button
                type="button"
                onClick={loadEmployeeDailyStatus}
                className="px-3.5 py-2 bg-white hover:bg-slate-50 border border-slate-200 rounded-2xl text-xs font-bold text-slate-700 flex items-center gap-1.5 shadow-sm transition-all active:scale-95 cursor-pointer"
                title="Segarkan Data Status"
              >
                <i className={`fa-solid fa-rotate text-brand-600 ${teamLoading ? 'fa-spin' : ''}`}></i>
                <span>Segarkan</span>
              </button>
            </div>

            {/* Main Card */}
            <div className="bg-white rounded-3xl p-5 border border-slate-200 shadow-sm space-y-4">
              <div className="flex items-center gap-2.5 border-b border-slate-100 pb-3">
                <div className="w-9 h-9 rounded-2xl bg-brand-50 text-brand-600 flex items-center justify-center text-sm shadow-inner">
                  <i className="fa-solid fa-users-viewfinder"></i>
                </div>
                <div>
                  <h3 className="font-extrabold text-sm text-slate-900 leading-tight">Status Kehadiran Hari Ini</h3>
                  <p className="text-[11px] text-slate-500 font-medium">Data real-time rekan kerja hari ini</p>
                </div>
              </div>

              {/* Filter Pills */}
              <div className="flex items-center gap-1.5 overflow-x-auto pb-1 hide-scrollbar">
                {(
                  [
                    { key: 'masuk', label: 'Sudah Masuk' },
                    { key: 'belum_masuk', label: 'Belum Masuk' },
                    { key: 'izin', label: 'Izin' },
                    { key: 'libur', label: 'Libur' },
                    { key: 'cuti', label: 'Cuti' },
                  ] as const
                ).map(({ key, label }) => {
                  const isActive = activeTeamFilter === key;
                  return (
                    <button
                      key={key}
                      type="button"
                      onClick={() => setActiveTeamFilter(key)}
                      className={`px-3 py-1.5 rounded-xl font-bold shrink-0 transition-all text-[11px] flex items-center gap-1.5 cursor-pointer ${
                        isActive
                          ? 'bg-brand-600 text-white shadow-sm'
                          : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
                      }`}
                    >
                      <span>{label}</span>
                      <span
                        className={`px-1.5 py-0.5 rounded-full text-[10px] font-black leading-none ${
                          isActive ? 'bg-white/20 text-white' : 'bg-slate-200 text-slate-700'
                        }`}
                      >
                        {teamCounters[key]}
                      </span>
                    </button>
                  );
                })}
              </div>

              {/* Team list */}
              <div className="space-y-2 max-h-[58vh] overflow-y-auto pr-1">
                {teamLoading ? (
                  <div className="text-xs text-slate-400 text-center py-6">Memuat status karyawan...</div>
                ) : teamDailyData.filter(i => i.category === activeTeamFilter).length === 0 ? (
                  <div className="text-xs text-slate-400 text-center py-6">Tidak ada karyawan pada status ini.</div>
                ) : (
                  teamDailyData
                    .filter(i => i.category === activeTeamFilter)
                    .map(item => (
                      <div
                        key={item.id}
                        className="flex items-center justify-between p-2.5 rounded-2xl bg-slate-50 border border-slate-100 hover:bg-slate-100/70 transition-colors"
                      >
                        <div className="flex items-center gap-2.5 min-w-0">
                          <div className="w-8 h-8 rounded-xl bg-brand-50 text-brand-600 flex items-center justify-center text-xs font-black shrink-0 border border-brand-100">
                            {item.nama.charAt(0).toUpperCase()}
                          </div>
                          <div className="min-w-0">
                            <div className="font-bold text-xs text-slate-900 truncate">{item.nama}</div>
                            <div className="text-[10px] text-slate-500 font-semibold truncate">{item.divisi}</div>
                          </div>
                        </div>
                        <div className="shrink-0 ml-2">
                          <span className={`px-2.5 py-1 rounded-full text-[10px] font-extrabold border ${item.statusColor}`}>
                            {item.statusText}
                          </span>
                        </div>
                      </div>
                    ))
                )}
              </div>
            </div>
          </section>
        )}

        {/* VIEW 3: DASHBOARD ADMIN */}
        {currentView === 'dashboard' && (
          <section className="flex-1 w-full max-w-7xl mx-auto p-4 sm:p-6 flex flex-col gap-6">
            {!currentAdmin ? (
              /* Admin Login Form */
              <div className="max-w-md w-full mx-auto bg-white p-6 sm:p-8 rounded-3xl shadow-xl border border-slate-100 my-auto">
                <div className="text-center mb-6">
                  <div className="w-16 h-16 rounded-2xl bg-indigo-50 text-indigo-600 flex items-center justify-center mx-auto mb-3 shadow-inner">
                    <i className="fa-solid fa-shield-halved text-2xl"></i>
                  </div>
                  <h2 className="text-2xl font-extrabold text-slate-900">Admin Dashboard</h2>
                  <p className="text-xs text-slate-500 mt-1">Masuk untuk mengelola data master dan presensi</p>
                </div>

                <form onSubmit={handleAdminLogin} className="space-y-4">
                  <div>
                    <label className="block text-xs font-bold text-slate-600 mb-1.5 uppercase tracking-wider">Username / Email</label>
                    <div className="relative">
                      <span className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none text-slate-400">
                        <i className="fa-solid fa-user-shield text-sm"></i>
                      </span>
                      <input
                        type="text"
                        value={adminLoginId}
                        onChange={e => setAdminLoginId(e.target.value)}
                        required
                        placeholder="admin"
                        className="w-full pl-10 pr-4 py-3 bg-slate-50 border border-slate-200 rounded-2xl text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 font-medium transition-all"
                      />
                    </div>
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-600 mb-1.5 uppercase tracking-wider">Password</label>
                    <div className="relative">
                      <span className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none text-slate-400">
                        <i className="fa-solid fa-lock text-sm"></i>
                      </span>
                      <input
                        type="password"
                        value={adminLoginPass}
                        onChange={e => setAdminLoginPass(e.target.value)}
                        required
                        placeholder="••••••••"
                        className="w-full pl-10 pr-4 py-3 bg-slate-50 border border-slate-200 rounded-2xl text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 font-medium transition-all"
                      />
                    </div>
                  </div>

                  <button
                    type="submit"
                    disabled={adminLoginLoading}
                    className="w-full py-3.5 px-4 bg-slate-900 hover:bg-black active:scale-[0.99] text-white font-bold rounded-2xl shadow-lg shadow-slate-900/20 transition-all flex items-center justify-center gap-2 cursor-pointer disabled:opacity-60"
                  >
                    {adminLoginLoading ? (
                      <>
                        <i className="fa-solid fa-spinner fa-spin"></i>
                        <span>Memeriksa...</span>
                      </>
                    ) : (
                      <span>Masuk Dashboard</span>
                    )}
                  </button>
                </form>
              </div>
            ) : (
              /* Admin App Panel */
              <div className="flex flex-col md:flex-row gap-6 w-full items-start">
                {/* Sidebar */}
                <aside
                  onMouseEnter={() => {
                    if (isSidebarCollapsed && window.innerWidth >= 768) {
                      setIsSidebarCollapsed(false);
                    }
                  }}
                  onMouseLeave={() => {
                    if (localStorage.getItem('presensi_admin_sidebar_collapsed') === 'true' && window.innerWidth >= 768) {
                      setIsSidebarCollapsed(true);
                    }
                  }}
                  className={`w-full ${
                    isSidebarCollapsed ? 'md:w-20' : 'md:w-64'
                  } bg-white rounded-3xl p-4 border border-slate-200 shadow-sm flex flex-col gap-2 shrink-0 transition-all duration-300`}
                >
                  <div className="flex items-center gap-3 p-2 mb-2 border-b border-slate-100">
                    <div className="w-10 h-10 rounded-2xl bg-indigo-50 text-indigo-600 flex items-center justify-center font-black text-sm shrink-0">
                      <i className="fa-solid fa-user-shield"></i>
                    </div>
                    {!isSidebarCollapsed && (
                      <div className="min-w-0">
                        <h4 className="font-extrabold text-xs text-slate-900 truncate">{currentAdmin.nama_lengkap}</h4>
                        <span className="text-[10px] text-slate-400 font-semibold uppercase">{currentAdmin.role || 'Super Admin'}</span>
                      </div>
                    )}
                  </div>

                  <nav className="flex flex-col gap-1 w-full">
                    {[
                      { key: 'rekap', icon: 'fa-chart-pie', label: 'Rekapan Presensi' },
                      { key: 'karyawan', icon: 'fa-users', label: 'Master Karyawan' },
                      { key: 'lokasi', icon: 'fa-location-dot', label: 'Master Lokasi' },
                      { key: 'shift', icon: 'fa-business-time', label: 'Master Shift & Jam' },
                      { key: 'divisi', icon: 'fa-sitemap', label: 'Master Divisi' },
                      { key: 'lembur', icon: 'fa-clock-rotate-left', label: 'Pengajuan Lembur' },
                      { key: 'jatah', icon: 'fa-calendar-check', label: 'Jatah Libur & Cuti' },
                      { key: 'pengumuman', icon: 'fa-bullhorn', label: 'Siaran Pengumuman' },
                      { key: 'useradmin', icon: 'fa-user-gear', label: 'Master Admin' },
                    ].map(({ key, icon, label }) => {
                      const isActive = activeAdminTab === key;
                      return (
                        <button
                          key={key}
                          onClick={() => switchAdminTab(key as any)}
                          className={`flex items-center gap-3 px-3 py-2.5 rounded-2xl text-xs font-bold transition-all text-left cursor-pointer ${
                            isActive ? 'bg-brand-50 text-brand-700' : 'text-slate-600 hover:bg-slate-50'
                          }`}
                        >
                          <i className={`fa-solid ${icon} w-5 text-center text-sm`}></i>
                          {!isSidebarCollapsed && <span>{label}</span>}
                        </button>
                      );
                    })}
                  </nav>

                  <button
                    onClick={logoutAdmin}
                    className="mt-4 flex items-center gap-3 px-3 py-2.5 rounded-2xl text-xs font-bold text-rose-600 hover:bg-rose-50 transition-all text-left border-t border-slate-100 pt-3 cursor-pointer"
                  >
                    <i className="fa-solid fa-arrow-right-from-bracket w-5 text-center text-sm"></i>
                    {!isSidebarCollapsed && <span>Keluar Admin</span>}
                  </button>
                </aside>

                {/* Right Content */}
                <div className="flex-1 w-full bg-white rounded-3xl p-5 sm:p-6 border border-slate-200 shadow-sm min-w-0">
                  {/* 1. PANEL REKAPAN PRESENSI */}
                  {activeAdminTab === 'rekap' && (
                    <div className="space-y-4">
                      <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-3 border-b border-slate-100 pb-4">
                        <div>
                          <h3 className="font-extrabold text-base text-slate-900">Rekapan Presensi Harian</h3>
                          <p className="text-xs text-slate-500">Data presensi real-time seluruh karyawan</p>
                        </div>
                        <div className="flex flex-wrap items-center gap-2">
                          {/* Penggeser Tanggal Kiri & Kanan */}
                          <div className="flex items-center gap-1 bg-slate-100/90 border border-slate-200 p-1 rounded-2xl shadow-xs">
                            <button
                              type="button"
                              onClick={() => shiftRekapTanggal(-1)}
                              className="w-8 h-8 rounded-xl bg-white hover:bg-slate-50 active:scale-95 text-slate-700 flex items-center justify-center transition-all shadow-xs cursor-pointer border border-slate-200/80"
                              title="Geser Hari Sebelumnya"
                            >
                              <i className="fa-solid fa-chevron-left text-xs"></i>
                            </button>

                            <div className="relative flex items-center gap-2 px-3 py-1.5 bg-white rounded-xl border border-slate-200/90 shadow-xs hover:border-brand-400 transition-all cursor-pointer group">
                              <i className="fa-regular fa-calendar text-brand-600 text-xs group-hover:scale-110 transition-transform"></i>
                              <span className="text-xs font-bold text-slate-800 select-none">
                                {formatIndoDate(filterRekapTanggal) || 'Pilih Tanggal'}
                              </span>
                              <i className="fa-solid fa-chevron-down text-[9px] text-slate-400 ml-0.5"></i>
                              <input
                                type="date"
                                value={filterRekapTanggal}
                                onChange={e => {
                                  setFilterRekapTanggal(e.target.value);
                                  loadAdminRekapan(e.target.value);
                                }}
                                className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                                title="Klik untuk memilih tanggal"
                              />
                            </div>

                            <button
                              type="button"
                              onClick={() => shiftRekapTanggal(1)}
                              className="w-8 h-8 rounded-xl bg-white hover:bg-slate-50 active:scale-95 text-slate-700 flex items-center justify-center transition-all shadow-xs cursor-pointer border border-slate-200/80"
                              title="Geser Hari Berikutnya"
                            >
                              <i className="fa-solid fa-chevron-right text-xs"></i>
                            </button>
                          </div>

                          {filterRekapTanggal !== getLocalISODate() && (
                            <button
                              type="button"
                              onClick={() => {
                                const today = getLocalISODate();
                                setFilterRekapTanggal(today);
                                loadAdminRekapan(today);
                              }}
                              className="px-3 py-2 text-xs font-bold bg-white hover:bg-slate-50 text-brand-600 rounded-xl border border-slate-200 shadow-xs transition-all active:scale-95 cursor-pointer"
                            >
                              Hari Ini
                            </button>
                          )}

                          {/* Tombol Export Excel Periode */}
                          <button
                            type="button"
                            onClick={openExportExcelModal}
                            className="px-3.5 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-sm transition-all cursor-pointer active:scale-95"
                          >
                            <i className="fa-solid fa-file-excel text-sm"></i>
                            <span>Export Excel</span>
                          </button>
                        </div>
                      </div>

                      {/* Kartu Statistik 5 Status Sekaligus Filter Interaktif */}
                      <div className="space-y-2">
                        <div className="grid grid-cols-2 sm:grid-cols-5 gap-2.5">
                          {/* 1. Masuk */}
                          <button
                            type="button"
                            onClick={() => setRekapFilterCategory(prev => prev === 'masuk' ? 'all' : 'masuk')}
                            className={`p-3 rounded-2xl text-left border transition-all cursor-pointer active:scale-95 ${
                              rekapFilterCategory === 'masuk'
                                ? 'bg-emerald-600 text-white border-emerald-600 shadow-md ring-2 ring-emerald-500/30'
                                : 'bg-emerald-50 hover:bg-emerald-100/70 border-emerald-100 text-emerald-800'
                            }`}
                          >
                            <div className="flex items-center justify-between">
                              <span className={`text-[10px] font-extrabold uppercase ${rekapFilterCategory === 'masuk' ? 'text-emerald-100' : 'text-emerald-600'}`}>
                                Masuk
                              </span>
                              <i className="fa-solid fa-arrow-right-to-bracket text-xs opacity-80"></i>
                            </div>
                            <div className={`text-2xl font-black mt-1 ${rekapFilterCategory === 'masuk' ? 'text-white' : 'text-emerald-700'}`}>
                              {rekapStats.masuk}
                            </div>
                            <span className="text-[10px] font-semibold opacity-80 mt-0.5 block">
                              {rekapFilterCategory === 'masuk' ? '✓ Filter Aktif' : 'Klik untuk filter'}
                            </span>
                          </button>

                          {/* 2. Belum Masuk */}
                          <button
                            type="button"
                            onClick={() => setRekapFilterCategory(prev => prev === 'belum_masuk' ? 'all' : 'belum_masuk')}
                            className={`p-3 rounded-2xl text-left border transition-all cursor-pointer active:scale-95 ${
                              rekapFilterCategory === 'belum_masuk'
                                ? 'bg-slate-800 text-white border-slate-800 shadow-md ring-2 ring-slate-500/30'
                                : 'bg-slate-100 hover:bg-slate-200/80 border-slate-200 text-slate-800'
                            }`}
                          >
                            <div className="flex items-center justify-between">
                              <span className={`text-[10px] font-extrabold uppercase ${rekapFilterCategory === 'belum_masuk' ? 'text-slate-300' : 'text-slate-600'}`}>
                                Belum Masuk
                              </span>
                              <i className="fa-solid fa-user-clock text-xs opacity-80"></i>
                            </div>
                            <div className={`text-2xl font-black mt-1 ${rekapFilterCategory === 'belum_masuk' ? 'text-white' : 'text-slate-800'}`}>
                              {rekapStats.belum_masuk}
                            </div>
                            <span className="text-[10px] font-semibold opacity-80 mt-0.5 block">
                              {rekapFilterCategory === 'belum_masuk' ? '✓ Filter Aktif' : 'Klik untuk filter'}
                            </span>
                          </button>

                          {/* 3. Sakit */}
                          <button
                            type="button"
                            onClick={() => setRekapFilterCategory(prev => prev === 'sakit' ? 'all' : 'sakit')}
                            className={`p-3 rounded-2xl text-left border transition-all cursor-pointer active:scale-95 ${
                              rekapFilterCategory === 'sakit'
                                ? 'bg-amber-600 text-white border-amber-600 shadow-md ring-2 ring-amber-500/30'
                                : 'bg-amber-50 hover:bg-amber-100/70 border-amber-100 text-amber-800'
                            }`}
                          >
                            <div className="flex items-center justify-between">
                              <span className={`text-[10px] font-extrabold uppercase ${rekapFilterCategory === 'sakit' ? 'text-amber-100' : 'text-amber-600'}`}>
                                Sakit
                              </span>
                              <i className="fa-solid fa-briefcase-medical text-xs opacity-80"></i>
                            </div>
                            <div className={`text-2xl font-black mt-1 ${rekapFilterCategory === 'sakit' ? 'text-white' : 'text-amber-700'}`}>
                              {rekapStats.sakit}
                            </div>
                            <span className="text-[10px] font-semibold opacity-80 mt-0.5 block">
                              {rekapFilterCategory === 'sakit' ? '✓ Filter Aktif' : 'Klik untuk filter'}
                            </span>
                          </button>

                          {/* 4. Libur */}
                          <button
                            type="button"
                            onClick={() => setRekapFilterCategory(prev => prev === 'libur' ? 'all' : 'libur')}
                            className={`p-3 rounded-2xl text-left border transition-all cursor-pointer active:scale-95 ${
                              rekapFilterCategory === 'libur'
                                ? 'bg-sky-600 text-white border-sky-600 shadow-md ring-2 ring-sky-500/30'
                                : 'bg-sky-50 hover:bg-sky-100/70 border-sky-100 text-sky-800'
                            }`}
                          >
                            <div className="flex items-center justify-between">
                              <span className={`text-[10px] font-extrabold uppercase ${rekapFilterCategory === 'libur' ? 'text-sky-100' : 'text-sky-600'}`}>
                                Libur
                              </span>
                              <i className="fa-solid fa-umbrella-beach text-xs opacity-80"></i>
                            </div>
                            <div className={`text-2xl font-black mt-1 ${rekapFilterCategory === 'libur' ? 'text-white' : 'text-sky-700'}`}>
                              {rekapStats.libur}
                            </div>
                            <span className="text-[10px] font-semibold opacity-80 mt-0.5 block">
                              {rekapFilterCategory === 'libur' ? '✓ Filter Aktif' : 'Klik untuk filter'}
                            </span>
                          </button>

                          {/* 5. Cuti */}
                          <button
                            type="button"
                            onClick={() => setRekapFilterCategory(prev => prev === 'cuti' ? 'all' : 'cuti')}
                            className={`p-3 rounded-2xl text-left border transition-all cursor-pointer active:scale-95 ${
                              rekapFilterCategory === 'cuti'
                                ? 'bg-purple-600 text-white border-purple-600 shadow-md ring-2 ring-purple-500/30'
                                : 'bg-purple-50 hover:bg-purple-100/70 border-purple-100 text-purple-800'
                            }`}
                          >
                            <div className="flex items-center justify-between">
                              <span className={`text-[10px] font-extrabold uppercase ${rekapFilterCategory === 'cuti' ? 'text-purple-100' : 'text-purple-600'}`}>
                                Cuti
                              </span>
                              <i className="fa-solid fa-calendar-check text-xs opacity-80"></i>
                            </div>
                            <div className={`text-2xl font-black mt-1 ${rekapFilterCategory === 'cuti' ? 'text-white' : 'text-purple-700'}`}>
                              {rekapStats.cuti}
                            </div>
                            <span className="text-[10px] font-semibold opacity-80 mt-0.5 block">
                              {rekapFilterCategory === 'cuti' ? '✓ Filter Aktif' : 'Klik untuk filter'}
                            </span>
                          </button>
                        </div>

                        {/* Reset Filter Button jika filter aktif */}
                        {rekapFilterCategory !== 'all' && (
                          <div className="flex items-center justify-between bg-slate-50 border border-slate-200 px-3.5 py-2 rounded-xl text-xs">
                            <span className="font-semibold text-slate-700">
                              Filter aktif: <b className="text-brand-700 uppercase">{rekapFilterCategory.replace('_', ' ')}</b> ({filteredRekapList.length} karyawan ditemukan)
                            </span>
                            <button
                              type="button"
                              onClick={() => setRekapFilterCategory('all')}
                              className="font-bold text-brand-600 hover:text-brand-800 hover:underline cursor-pointer"
                            >
                              Tampilkan Semua (Reset Filter)
                            </button>
                          </div>
                        )}
                      </div>

                      {/* Search Bar Rekapan */}
                      <div className="relative">
                        <span className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none text-slate-400">
                          <i className="fa-solid fa-magnifying-glass text-xs"></i>
                        </span>
                        <input
                          type="text"
                          value={searchRekapQuery}
                          onChange={e => setSearchRekapQuery(e.target.value)}
                          placeholder="Cari karyawan berdasarkan nama atau divisi..."
                          className="w-full pl-9 pr-8 py-2.5 bg-slate-50 border border-slate-200 rounded-2xl text-xs font-medium focus:outline-none focus:ring-2 focus:ring-brand-500/20 focus:border-brand-500 transition-all placeholder:text-slate-400"
                        />
                        {searchRekapQuery && (
                          <button
                            type="button"
                            onClick={() => setSearchRekapQuery('')}
                            className="absolute inset-y-0 right-0 pr-3 flex items-center text-slate-400 hover:text-slate-600 cursor-pointer"
                          >
                            <i className="fa-solid fa-xmark text-xs"></i>
                          </button>
                        )}
                      </div>

                      {/* Table Rekap */}
                      <div className="overflow-x-auto border border-slate-200 rounded-2xl">
                        <table className="w-full text-xs text-left text-slate-600">
                          <thead className="bg-slate-50 text-slate-700 uppercase text-[10px] font-black border-b border-slate-200">
                            <tr>
                              <th className="p-3">Karyawan</th>
                              <th className="p-3">Divisi / Shift</th>
                              <th className="p-3">Jam Masuk</th>
                              <th className="p-3">Jam Keluar</th>
                              <th className="p-3">Status</th>
                              <th className="p-3">Lokasi</th>
                              <th className="p-3">Keterangan</th>
                              <th className="p-3 min-w-[130px]">Foto Presensi</th>
                              <th className="p-3 text-right">Aksi</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-100 font-medium">
                            {filteredRekapList.length === 0 ? (
                              <tr>
                                <td colSpan={9} className="p-4 text-center text-slate-400">
                                  {searchRekapQuery || rekapFilterCategory !== 'all' ? 'Tidak ada data karyawan yang sesuai filter.' : 'Tidak ada riwayat presensi pada tanggal ini.'}
                                </td>
                              </tr>
                            ) : (
                              filteredRekapList.map((item, index) => {
                                let statusBadgeClass = 'bg-slate-100 text-slate-700';
                                let statusText = 'Hadir';

                                if (item.category === 'belum_masuk') {
                                  statusBadgeClass = 'bg-slate-100 text-slate-600 border border-slate-200';
                                  statusText = 'Belum Masuk';
                                } else {
                                  const rawStatus = (item.keluar?.status || item.masuk?.status || 'Hadir').toUpperCase();
                                  if (rawStatus === 'TERLAMBAT') {
                                    statusBadgeClass = 'bg-amber-100 text-amber-700 border border-amber-200';
                                    statusText = 'Terlambat';
                                  } else if (rawStatus === 'SAKIT') {
                                    statusBadgeClass = 'bg-amber-100 text-amber-800 border border-amber-200';
                                    statusText = 'Sakit';
                                  } else if (rawStatus === 'LIBUR') {
                                    statusBadgeClass = 'bg-sky-100 text-sky-700 border border-sky-200';
                                    statusText = 'Libur';
                                  } else if (rawStatus === 'CUTI') {
                                    statusBadgeClass = 'bg-purple-100 text-purple-700 border border-purple-200';
                                    statusText = 'Cuti';
                                  } else if (rawStatus === 'IZIN') {
                                    statusBadgeClass = 'bg-orange-100 text-orange-700 border border-orange-200';
                                    statusText = 'Izin';
                                  } else if (item.masuk && item.keluar) {
                                    statusBadgeClass = 'bg-indigo-100 text-indigo-700 border border-indigo-200';
                                    statusText = 'Selesai (Pulang)';
                                  } else if (item.masuk) {
                                    statusBadgeClass = 'bg-emerald-100 text-emerald-700 border border-emerald-200';
                                    statusText = 'Hadir (Masih Kerja)';
                                  }
                                }

                                const jamMasukStr = item.masuk ? formatTime24(item.masuk.created_at) : '-';
                                const jamKeluarStr = item.keluar ? formatTime24(item.keluar.created_at) : '-';
                                const ketStr = [item.masuk?.keterangan, item.keluar?.keterangan].filter(Boolean).join(' | ') || '-';

                                return (
                                  <tr key={item.karyawan_id || item.karyawan_nama} className="hover:bg-slate-50 transition-colors">
                                    <td className="p-3">
                                      <div className="flex items-center gap-2">
                                        <span className="inline-flex items-center justify-center min-w-5 h-5 px-1.5 rounded-md bg-slate-100 text-slate-700 text-[10px] font-extrabold border border-slate-200 shrink-0">
                                          {index + 1}
                                        </span>
                                        <div className="font-bold text-slate-900">{item.karyawan_nama}</div>
                                      </div>
                                    </td>
                                    <td className="p-3">
                                      <div className="text-slate-700 font-medium">{item.divisi || '-'}</div>
                                      <div className="text-[11px] font-bold text-brand-600">{item.shift_nama || '-'}</div>
                                    </td>
                                    <td className="p-3">
                                      <div className={`font-mono font-bold ${item.masuk ? 'text-emerald-700' : 'text-slate-400'}`}>
                                        {jamMasukStr}
                                      </div>
                                      {item.masuk?.status === 'Terlambat' && (
                                        <span className="text-[10px] font-bold text-amber-600 block">Terlambat</span>
                                      )}
                                    </td>
                                    <td className="p-3">
                                      <div className={`font-mono font-bold ${item.keluar ? 'text-rose-700' : 'text-slate-400'}`}>
                                        {jamKeluarStr}
                                      </div>
                                      {item.keluar && (
                                        <span className="text-[10px] font-bold text-slate-500 block uppercase">
                                          {item.keluar.status}
                                        </span>
                                      )}
                                    </td>
                                    <td className="p-3">
                                      <span className={`px-2.5 py-1 rounded-full text-[10px] font-extrabold ${statusBadgeClass}`}>
                                        {statusText}
                                      </span>
                                    </td>
                                    <td className="p-3 text-slate-600 text-xs">{item.lokasi_nama || '-'}</td>
                                    <td className="p-3 text-slate-500 max-w-xs truncate" title={ketStr}>
                                      {ketStr}
                                    </td>
                                    <td className="p-3">
                                      {item.masuk?.foto_url || item.keluar?.foto_url ? (
                                        <div className="flex items-center gap-2">
                                          {item.masuk?.foto_url && (
                                            <div className="flex flex-col items-center gap-1">
                                              <div
                                                onClick={() => openPhotoPreviewModal(item.masuk!.foto_url!, item.karyawan_nama, `Foto Masuk (${jamMasukStr})`)}
                                                className="group relative w-12 h-12 rounded-xl overflow-hidden bg-slate-900 border border-emerald-200 shadow-sm cursor-pointer hover:ring-2 hover:ring-emerald-500 transition-all shrink-0"
                                                title="Klik untuk memperbesar Foto Masuk"
                                              >
                                                <img
                                                  src={formatDirectImageUrl(item.masuk.foto_url)}
                                                  alt={`Masuk ${item.karyawan_nama}`}
                                                  className="w-full h-full object-cover group-hover:scale-110 transition-transform"
                                                  onError={(e) => {
                                                    (e.currentTarget as HTMLElement).style.display = 'none';
                                                    (e.currentTarget.parentElement?.querySelector('.fallback-masuk') as HTMLElement | null)?.classList.remove('hidden');
                                                  }}
                                                />
                                                <div className="fallback-masuk hidden absolute inset-0 bg-emerald-50 text-emerald-700 flex flex-col items-center justify-center text-[10px] font-bold">
                                                  <i className="fa-solid fa-camera text-xs"></i>
                                                </div>
                                                <div className="absolute inset-0 bg-black/0 group-hover:bg-black/30 flex items-center justify-center transition-all opacity-0 group-hover:opacity-100 text-white text-[10px]">
                                                  <i className="fa-solid fa-expand"></i>
                                                </div>
                                              </div>
                                              <span className="text-[9px] font-extrabold text-emerald-700 bg-emerald-50 px-1 rounded border border-emerald-100">
                                                Masuk
                                              </span>
                                            </div>
                                          )}
                                          {item.keluar?.foto_url && (
                                            <div className="flex flex-col items-center gap-1">
                                              <div
                                                onClick={() => openPhotoPreviewModal(item.keluar!.foto_url!, item.karyawan_nama, `Foto Keluar (${jamKeluarStr})`)}
                                                className="group relative w-12 h-12 rounded-xl overflow-hidden bg-slate-900 border border-rose-200 shadow-sm cursor-pointer hover:ring-2 hover:ring-rose-500 transition-all shrink-0"
                                                title="Klik untuk memperbesar Foto Keluar"
                                              >
                                                <img
                                                  src={formatDirectImageUrl(item.keluar.foto_url)}
                                                  alt={`Keluar ${item.karyawan_nama}`}
                                                  className="w-full h-full object-cover group-hover:scale-110 transition-transform"
                                                  onError={(e) => {
                                                    (e.currentTarget as HTMLElement).style.display = 'none';
                                                    (e.currentTarget.parentElement?.querySelector('.fallback-keluar') as HTMLElement | null)?.classList.remove('hidden');
                                                  }}
                                                />
                                                <div className="fallback-keluar hidden absolute inset-0 bg-rose-50 text-rose-700 flex flex-col items-center justify-center text-[10px] font-bold">
                                                  <i className="fa-solid fa-camera text-xs"></i>
                                                </div>
                                                <div className="absolute inset-0 bg-black/0 group-hover:bg-black/30 flex items-center justify-center transition-all opacity-0 group-hover:opacity-100 text-white text-[10px]">
                                                  <i className="fa-solid fa-expand"></i>
                                                </div>
                                              </div>
                                              <span className="text-[9px] font-extrabold text-rose-700 bg-rose-50 px-1 rounded border border-rose-100">
                                                Keluar
                                              </span>
                                            </div>
                                          )}
                                        </div>
                                      ) : (
                                        <span className="text-slate-400 text-[11px] italic">Tanpa foto</span>
                                      )}
                                    </td>
                                    <td className="p-3 text-right whitespace-nowrap space-x-1">
                                      <button
                                        type="button"
                                        onClick={() => {
                                          setEditingRekapItem(item);
                                          setEditRekapShift(item.shift_nama || 'Reguler');
                                          if (item.masuk) {
                                            const d = new Date(item.masuk.created_at);
                                            setEditRekapJamMasuk(`${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`);
                                            setEditRekapStatusMasuk(item.masuk.status === 'Terlambat' ? 'Terlambat' : 'Hadir');
                                          } else {
                                            setEditRekapJamMasuk('');
                                            setEditRekapStatusMasuk('Belum Masuk');
                                          }

                                          if (item.keluar) {
                                            const d = new Date(item.keluar.created_at);
                                            setEditRekapJamKeluar(`${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`);
                                            setEditRekapStatusKeluar((item.keluar.status || 'PULANG').toUpperCase());
                                          } else {
                                            setEditRekapJamKeluar('');
                                            setEditRekapStatusKeluar('Belum Keluar');
                                          }

                                          setEditRekapKeterangan([item.masuk?.keterangan, item.keluar?.keterangan].filter(Boolean).join(' | ') || '');
                                          setModalEditRekapOpen(true);
                                        }}
                                        className="px-2.5 py-1 rounded-lg bg-indigo-50 text-indigo-600 hover:bg-indigo-100 font-bold transition-all active:scale-95 cursor-pointer"
                                        title="Edit Presensi"
                                      >
                                        <i className="fa-solid fa-pen"></i>
                                      </button>
                                      <button
                                        type="button"
                                        onClick={() => {
                                          const ids = item.records.map(r => r.id);
                                          setDeleteConfirmMsg(`Apakah Anda yakin ingin menghapus seluruh data presensi ${item.karyawan_nama} pada tanggal ini?`);
                                          setDeleteActionFn(() => async () => {
                                            const { error } = await supabaseClient.from('presensi_presensi').delete().in('id', ids);
                                            if (error) throw error;
                                            showToast(`Presensi ${item.karyawan_nama} berhasil dihapus!`, 'success');
                                            loadAdminRekapan();
                                          });
                                          setModalConfirmDeleteOpen(true);
                                        }}
                                        className="px-2.5 py-1 rounded-lg bg-rose-50 text-rose-600 hover:bg-rose-100 font-bold transition-all active:scale-95 cursor-pointer"
                                        title="Hapus Presensi"
                                      >
                                        <i className="fa-solid fa-trash"></i>
                                      </button>
                                    </td>
                                  </tr>
                                );
                              })
                            )}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  )}

                  {/* 2. PANEL MASTER KARYAWAN */}
                  {activeAdminTab === 'karyawan' && (
                    <div className="space-y-4">
                      <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                        <div>
                          <h3 className="font-extrabold text-base text-slate-900">Master Data Karyawan</h3>
                          <p className="text-xs text-slate-500">Kelola akun, tanggal lahir, dan divisi karyawan</p>
                        </div>
                        <button
                          onClick={async () => {
                            setEditingKaryawanId(null);
                            setKaryawanFormNama('');
                            setKaryawanFormTglLahir('');
                            setKaryawanFormUsername('');
                            setKaryawanFormWa('');
                            setKaryawanFormEmail('');
                            setKaryawanFormPass('');
                            setKaryawanFormDivisi(adminDivisiList[0]?.nama_divisi || 'Umum');
                            setKaryawanFormJalan('');
                            setSelectedProvId('');
                            setSelectedKotaId('');
                            setSelectedKecId('');
                            setSelectedDesaId('');
                            setKotaList([]);
                            setKecamatanList([]);
                            setDesaList([]);
                            await fetchProvinces();
                            setModalKaryawanOpen(true);
                          }}
                          className="px-3.5 py-2 bg-brand-600 hover:bg-brand-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-md cursor-pointer"
                        >
                          <i className="fa-solid fa-plus"></i> Tambah Karyawan
                        </button>
                      </div>

                      {/* Search Bar Master Karyawan */}
                      <div className="relative">
                        <span className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none text-slate-400">
                          <i className="fa-solid fa-magnifying-glass text-xs"></i>
                        </span>
                        <input
                          type="text"
                          value={searchKaryawanQuery}
                          onChange={e => setSearchKaryawanQuery(e.target.value)}
                          placeholder="Cari karyawan berdasarkan nama, username, atau divisi..."
                          className="w-full pl-9 pr-8 py-2.5 bg-slate-50 border border-slate-200 rounded-2xl text-xs font-medium focus:outline-none focus:ring-2 focus:ring-brand-500/20 focus:border-brand-500 transition-all placeholder:text-slate-400"
                        />
                        {searchKaryawanQuery && (
                          <button
                            type="button"
                            onClick={() => setSearchKaryawanQuery('')}
                            className="absolute inset-y-0 right-0 pr-3 flex items-center text-slate-400 hover:text-slate-600 cursor-pointer"
                          >
                            <i className="fa-solid fa-xmark text-xs"></i>
                          </button>
                        )}
                      </div>

                      <div className="overflow-x-auto border border-slate-200 rounded-2xl">
                        <table className="w-full text-xs text-left text-slate-600">
                          <thead className="bg-slate-50 text-slate-700 uppercase text-[10px] font-black border-b border-slate-200">
                            <tr>
                              <th className="p-3">Karyawan</th>
                              <th className="p-3">Kontak & Akun</th>
                              <th className="p-3">Divisi & Shift</th>
                              <th className="p-3">Sisa Kuota</th>
                              <th className="p-3 text-right">Aksi</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-100 font-medium">
                            {filteredAdminKaryawanList.length === 0 ? (
                              <tr>
                                <td colSpan={5} className="p-4 text-center text-slate-400">
                                  {searchKaryawanQuery ? 'Karyawan tidak ditemukan.' : 'Belum ada data master karyawan.'}
                                </td>
                              </tr>
                            ) : (
                              filteredAdminKaryawanList.map((k, index) => (
                                <tr key={k.id} className="hover:bg-slate-50">
                                  <td className="p-3">
                                    <div className="flex items-start gap-2">
                                      <span className="inline-flex items-center justify-center min-w-5 h-5 px-1.5 rounded-md bg-slate-100 text-slate-700 text-[10px] font-extrabold border border-slate-200 shrink-0 mt-0.5">
                                        {index + 1}
                                      </span>
                                      <div>
                                        <div className="font-bold text-slate-900">{k.nama_lengkap}</div>
                                        <div className="text-[10px] text-slate-400">Lahir: {k.tgl_lahir ? isoToDdmmyyyy(k.tgl_lahir) : '-'}</div>
                                        {k.alamat && (
                                          <div className="text-[10px] text-slate-500 font-medium truncate max-w-[220px] mt-0.5" title={getDisplayAddress(k.alamat)}>
                                            <i className="fa-solid fa-location-dot text-slate-400 mr-1"></i>
                                            {getDisplayAddress(k.alamat)}
                                          </div>
                                        )}
                                      </div>
                                    </div>
                                  </td>
                                  <td className="p-3">
                                    <div className="font-mono text-xs text-brand-600">{k.username}</div>
                                    <div className="text-[10px] text-slate-400">{k.no_wa || '-'} • {k.email || '-'}</div>
                                  </td>
                                  <td className="p-3">
                                    <span className="px-2 py-0.5 rounded-md bg-slate-100 text-slate-700 font-bold">{k.divisi || 'Umum'}</span>
                                    <div className="text-[10px] text-slate-400 mt-0.5">{(k.shifts || []).join(', ')}</div>
                                  </td>
                                  <td className="p-3">
                                    <span className="text-[11px] font-bold text-slate-700">Libur: {k.sisa_libur ?? 4} | Cuti: {k.sisa_cuti ?? 12}</span>
                                  </td>
                                  <td className="p-3 text-right space-x-1">
                                    <button
                                      onClick={async () => {
                                        setEditingKaryawanId(k.id);
                                        setKaryawanFormNama(k.nama_lengkap);
                                        setKaryawanFormTglLahir(isoToDdmmyyyy(k.tgl_lahir));
                                        setKaryawanFormUsername(k.username);
                                        setKaryawanFormWa(k.no_wa);
                                        setKaryawanFormEmail(k.email);
                                        setKaryawanFormPass('');
                                        setKaryawanFormDivisi(k.divisi || 'Umum');
                                        const addr = parseAddressData(k.alamat);
                                        setKaryawanFormJalan(addr.jalan || '');
                                        setModalKaryawanOpen(true);
                                        await openKaryawanRegionCascade(addr);
                                      }}
                                      className="px-2 py-1 rounded-lg bg-indigo-50 text-indigo-600 hover:bg-indigo-100 font-bold cursor-pointer"
                                      title="Edit Karyawan"
                                    >
                                      <i className="fa-solid fa-pen"></i>
                                    </button>
                                    <button
                                      onClick={() => {
                                        setDeleteConfirmMsg(`Apakah Anda yakin ingin menghapus data karyawan ${k.nama_lengkap}?`);
                                        setDeleteActionFn(() => async () => {
                                          await supabaseClient.from('karyawan_presensi').delete().eq('id', k.id);
                                          showToast('Data karyawan berhasil dihapus!', 'success');
                                          loadAdminKaryawan();
                                        });
                                        setModalConfirmDeleteOpen(true);
                                      }}
                                      className="px-2 py-1 rounded-lg bg-rose-50 text-rose-600 hover:bg-rose-100 font-bold cursor-pointer"
                                    >
                                      <i className="fa-solid fa-trash"></i>
                                    </button>
                                  </td>
                                </tr>
                              ))
                            )}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  )}

                  {/* 3. PANEL MASTER LOKASI */}
                  {activeAdminTab === 'lokasi' && (
                    <div className="space-y-4">
                      <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                        <div>
                          <h3 className="font-extrabold text-base text-slate-900">Master Lokasi Kantor</h3>
                          <p className="text-xs text-slate-500">Titik koordinat dan radius geofence presensi</p>
                        </div>
                        <button
                          onClick={() => {
                            setEditingLokasiId(null);
                            setLokasiNama('');
                            setLokasiLat('');
                            setLokasiLng('');
                            setLokasiRadius('100');
                            setModalLokasiOpen(true);
                          }}
                          className="px-3.5 py-2 bg-brand-600 hover:bg-brand-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-md cursor-pointer"
                        >
                          <i className="fa-solid fa-plus"></i> Tambah Lokasi
                        </button>
                      </div>

                      <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                        {adminLokasiList.length === 0 ? (
                          <div className="col-span-2 text-xs text-slate-400 text-center py-6">Belum ada lokasi kantor.</div>
                        ) : (
                          adminLokasiList.map(l => (
                            <div key={l.id} className="p-4 bg-slate-50 rounded-2xl border border-slate-200 flex flex-col justify-between gap-3">
                              <div>
                                <div className="flex items-center justify-between">
                                  <h4 className="font-extrabold text-sm text-slate-900">{l.nama_lokasi}</h4>
                                  <span className="px-2 py-0.5 rounded-full text-[10px] font-bold bg-indigo-100 text-indigo-700">Radius {l.radius_meter}m</span>
                                </div>
                                <div className="text-xs text-slate-500 font-mono mt-1">Lat: {l.latitude}, Lng: {l.longitude}</div>
                              </div>
                              <div className="flex items-center justify-end gap-1.5 pt-2 border-t border-slate-200">
                                <button
                                  onClick={() => {
                                    setEditingLokasiId(l.id);
                                    setLokasiNama(l.nama_lokasi);
                                    setLokasiLat(String(l.latitude));
                                    setLokasiLng(String(l.longitude));
                                    setLokasiRadius(String(l.radius_meter));
                                    setModalLokasiOpen(true);
                                  }}
                                  className="px-3 py-1 rounded-lg bg-indigo-100 text-indigo-700 hover:bg-indigo-200 text-xs font-bold cursor-pointer"
                                >
                                  <i className="fa-solid fa-pen mr-1"></i> Edit
                                </button>
                                <button
                                  onClick={() => {
                                    setDeleteConfirmMsg(`Hapus lokasi ${l.nama_lokasi}?`);
                                    setDeleteActionFn(() => async () => {
                                      await supabaseClient.from('lokasi_presensi').delete().eq('id', l.id);
                                      showToast('Lokasi berhasil dihapus!', 'success');
                                      loadAdminLokasi();
                                    });
                                    setModalConfirmDeleteOpen(true);
                                  }}
                                  className="px-3 py-1 rounded-lg bg-rose-100 text-rose-700 hover:bg-rose-200 text-xs font-bold cursor-pointer"
                                >
                                  <i className="fa-solid fa-trash mr-1"></i> Hapus
                                </button>
                              </div>
                            </div>
                          ))
                        )}
                      </div>
                    </div>
                  )}

                  {/* 4. PANEL MASTER SHIFT & JAM */}
                  {activeAdminTab === 'shift' && (
                    <div className="space-y-4">
                      <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                        <div>
                          <h3 className="font-extrabold text-base text-slate-900">Master Shift & Jam Kerja</h3>
                          <p className="text-xs text-slate-500">Pengaturan jadwal shift dan batas toleransi</p>
                        </div>
                        <button
                          onClick={() => {
                            setEditingShiftId(null);
                            setShiftNama('');
                            setShiftJamMasuk('');
                            setShiftJamPulang('');
                            setModalShiftOpen(true);
                          }}
                          className="px-3.5 py-2 bg-brand-600 hover:bg-brand-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-md cursor-pointer"
                        >
                          <i className="fa-solid fa-plus"></i> Tambah Shift
                        </button>
                      </div>

                      {/* Toleransi Form */}
                      <form
                        onSubmit={async e => {
                          e.preventDefault();
                          const { data: cur } = await supabaseClient.from('jam_kerja_presensi').select('id').limit(1).single();
                          if (cur) {
                            await supabaseClient.from('jam_kerja_presensi').update({ toleransi_menit: toleransiMenit }).eq('id', cur.id);
                          } else {
                            await supabaseClient.from('jam_kerja_presensi').insert([{ toleransi_menit: toleransiMenit }]);
                          }
                          showToast('Toleransi keterlambatan disimpan!', 'success');
                        }}
                        className="p-4 bg-slate-50 rounded-2xl border border-slate-200 flex flex-col sm:flex-row items-center gap-3"
                      >
                        <div className="flex items-center gap-2">
                          <i className="fa-solid fa-clock-rotate-left text-brand-600 text-base"></i>
                          <span className="text-xs font-bold text-slate-700">Toleransi Keterlambatan:</span>
                        </div>
                        <div className="flex items-center gap-2">
                          <input
                            type="number"
                            min="0"
                            max="120"
                            value={toleransiMenit}
                            onChange={e => setToleransiMenit(parseInt(e.target.value) || 0)}
                            className="w-20 px-3 py-1.5 bg-white border border-slate-200 rounded-xl text-xs font-bold text-center"
                          />
                          <span className="text-xs font-bold text-slate-500">Menit</span>
                          <button type="submit" className="px-4 py-1.5 bg-slate-900 hover:bg-black text-white rounded-xl text-xs font-bold cursor-pointer">
                            Simpan
                          </button>
                        </div>
                      </form>

                      {/* Shift Table */}
                      <div className="overflow-x-auto border border-slate-200 rounded-2xl">
                        <table className="w-full text-xs text-left text-slate-600">
                          <thead className="bg-slate-50 text-slate-700 uppercase text-[10px] font-black border-b border-slate-200">
                            <tr>
                              <th className="p-3">Nama Shift</th>
                              <th className="p-3">Jam Masuk (24 Jam)</th>
                              <th className="p-3">Jam Pulang (24 Jam)</th>
                              <th className="p-3 text-right">Aksi</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-100 font-medium">
                            {adminShiftList.length === 0 ? (
                              <tr>
                                <td colSpan={4} className="p-4 text-center text-slate-400">Belum ada master shift kerja.</td>
                              </tr>
                            ) : (
                              adminShiftList.map(s => (
                                <tr key={s.id} className="hover:bg-slate-50">
                                  <td className="p-3 font-bold text-slate-900">{s.nama_shift}</td>
                                  <td className="p-3 font-mono font-bold text-indigo-600">{s.jam_masuk?.substring(0, 5)}</td>
                                  <td className="p-3 font-mono font-bold text-indigo-600">{s.jam_pulang?.substring(0, 5)}</td>
                                  <td className="p-3 text-right space-x-1">
                                    <button
                                      onClick={() => {
                                        setEditingShiftId(s.id);
                                        setShiftNama(s.nama_shift);
                                        setShiftJamMasuk(s.jam_masuk?.substring(0, 5) || '');
                                        setShiftJamPulang(s.jam_pulang?.substring(0, 5) || '');
                                        setModalShiftOpen(true);
                                      }}
                                      className="px-2 py-1 rounded-lg bg-indigo-50 text-indigo-600 hover:bg-indigo-100 font-bold cursor-pointer"
                                    >
                                      <i className="fa-solid fa-pen"></i>
                                    </button>
                                    <button
                                      onClick={() => {
                                        setDeleteConfirmMsg(`Hapus shift ${s.nama_shift}?`);
                                        setDeleteActionFn(() => async () => {
                                          await supabaseClient.from('shift_presensi').delete().eq('id', s.id);
                                          showToast('Shift berhasil dihapus!', 'success');
                                          loadAdminShiftAndJam();
                                        });
                                        setModalConfirmDeleteOpen(true);
                                      }}
                                      className="px-2 py-1 rounded-lg bg-rose-50 text-rose-600 hover:bg-rose-100 font-bold cursor-pointer"
                                    >
                                      <i className="fa-solid fa-trash"></i>
                                    </button>
                                  </td>
                                </tr>
                              ))
                            )}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  )}

                  {/* 5. PANEL MASTER DIVISI */}
                  {activeAdminTab === 'divisi' && (
                    <div className="space-y-4">
                      <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                        <div>
                          <h3 className="font-extrabold text-base text-slate-900">Master Divisi</h3>
                          <p className="text-xs text-slate-500">Kelola divisi, hak cuti, dan shift terkait (Urutan A-Z)</p>
                        </div>
                        <button
                          onClick={async () => {
                            setEditingDivisiId(null);
                            setDivisiNama('');
                            setDivisiSelectedShifts([]);
                            setDivisiAllowCuti(true);
                            await loadAdminShiftAndJam();
                            setModalDivisiOpen(true);
                          }}
                          className="px-3.5 py-2 bg-brand-600 hover:bg-brand-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-md cursor-pointer"
                        >
                          <i className="fa-solid fa-plus"></i> Tambah Divisi
                        </button>
                      </div>

                      <div className="overflow-x-auto border border-slate-200 rounded-2xl">
                        <table className="w-full text-xs text-left text-slate-600">
                          <thead className="bg-slate-50 text-slate-700 uppercase text-[10px] font-black border-b border-slate-200">
                            <tr>
                              <th className="p-3">Nama Divisi</th>
                              <th className="p-3">Shift Terkait</th>
                              <th className="p-3">Hak Cuti</th>
                              <th className="p-3 text-right">Aksi</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-100 font-medium">
                            {adminDivisiList.length === 0 ? (
                              <tr>
                                <td colSpan={4} className="p-4 text-center text-slate-400">Belum ada divisi.</td>
                              </tr>
                            ) : (
                              adminDivisiList.map((d, index) => (
                                <tr key={d.id} className="hover:bg-slate-50">
                                  <td className="p-3">
                                    <div className="flex items-center gap-2">
                                      <span className="inline-flex items-center justify-center min-w-5 h-5 px-1.5 rounded-md bg-slate-100 text-slate-700 text-[10px] font-extrabold border border-slate-200 shrink-0">
                                        {index + 1}
                                      </span>
                                      <span className="font-bold text-slate-900">{d.nama_divisi}</span>
                                    </div>
                                  </td>
                                  <td className="p-3">
                                    <div className="flex flex-wrap gap-1">
                                      {(d.shifts || []).map(s => (
                                        <span key={s} className="px-2 py-0.5 rounded-lg bg-indigo-50 text-indigo-700 text-[11px] font-bold border border-indigo-100">
                                          {s}
                                        </span>
                                      ))}
                                    </div>
                                  </td>
                                  <td className="p-3">
                                    <span className={`px-2.5 py-1 rounded-full text-[10px] font-extrabold ${d.allow_cuti !== false ? 'bg-emerald-100 text-emerald-800 border border-emerald-200' : 'bg-slate-100 text-slate-500 border border-slate-200'}`}>
                                      <i className={`fa-solid ${d.allow_cuti !== false ? 'fa-check' : 'fa-xmark'} mr-1`}></i>
                                      {d.allow_cuti !== false ? 'Cuti Aktif' : 'Cuti Nonaktif'}
                                    </span>
                                  </td>
                                  <td className="p-3 text-right space-x-1">
                                    <button
                                      onClick={async () => {
                                        setEditingDivisiId(d.id);
                                        setDivisiNama(d.nama_divisi);
                                        setDivisiSelectedShifts(d.shifts || []);
                                        setDivisiAllowCuti(d.allow_cuti !== false);
                                        await loadAdminShiftAndJam();
                                        setModalDivisiOpen(true);
                                      }}
                                      className="px-2 py-1 rounded-lg bg-indigo-50 text-indigo-600 hover:bg-indigo-100 font-bold cursor-pointer"
                                      title="Edit"
                                    >
                                      <i className="fa-solid fa-pen"></i>
                                    </button>
                                    <button
                                      onClick={() => {
                                        setDeleteConfirmMsg(`Hapus divisi ${d.nama_divisi}?`);
                                        setDeleteActionFn(() => async () => {
                                          await supabaseClient.from('divisi_presensi').delete().eq('id', d.id);
                                          showToast('Divisi berhasil dihapus!', 'success');
                                          loadAdminDivisi();
                                        });
                                        setModalConfirmDeleteOpen(true);
                                      }}
                                      className="px-2 py-1 rounded-lg bg-rose-50 text-rose-600 hover:bg-rose-100 font-bold cursor-pointer"
                                      title="Hapus"
                                    >
                                      <i className="fa-solid fa-trash"></i>
                                    </button>
                                  </td>
                                </tr>
                              ))
                            )}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  )}

                  {/* 6. PANEL PENGAJUAN LEMBUR */}
                  {activeAdminTab === 'lembur' && (
                    <div className="space-y-4">
                      <div className="border-b border-slate-100 pb-3">
                        <h3 className="font-extrabold text-base text-slate-900">Pengajuan Lembur Karyawan</h3>
                        <p className="text-xs text-slate-500">Persetujuan atau penolakan pengajuan lembur</p>
                      </div>

                      <div className="overflow-x-auto border border-slate-200 rounded-2xl">
                        <table className="w-full text-xs text-left text-slate-600">
                          <thead className="bg-slate-50 text-slate-700 uppercase text-[10px] font-black border-b border-slate-200">
                            <tr>
                              <th className="p-3">Tanggal</th>
                              <th className="p-3">Karyawan</th>
                              <th className="p-3">Durasi</th>
                              <th className="p-3">Keterangan</th>
                              <th className="p-3">Bukti</th>
                              <th className="p-3">Status</th>
                              <th className="p-3 text-right">Aksi</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-100 font-medium">
                            {adminLemburList.length === 0 ? (
                              <tr>
                                <td colSpan={7} className="p-4 text-center text-slate-400">Belum ada pengajuan lembur.</td>
                              </tr>
                            ) : (
                              adminLemburList.map(l => (
                                <tr key={l.id} className="hover:bg-slate-50">
                                  <td className="p-3 font-mono text-xs">{l.tanggal}</td>
                                  <td className="p-3 font-bold text-slate-900">{l.karyawan_nama}</td>
                                  <td className="p-3 font-bold text-amber-700">{l.durasi_jam} Jam</td>
                                  <td className="p-3 text-slate-600 max-w-xs truncate">{l.keterangan}</td>
                                  <td className="p-3">
                                    {l.bukti_url ? (
                                      <button
                                        type="button"
                                        onClick={() => openPhotoPreviewModal(l.bukti_url!, l.karyawan_nama, `Bukti Lembur (${l.durasi_jam} Jam)`)}
                                        className="px-2.5 py-1 rounded-lg bg-indigo-50 text-indigo-600 hover:bg-indigo-100 font-bold inline-flex items-center gap-1 cursor-pointer"
                                      >
                                        <i className="fa-solid fa-image"></i> Lihat
                                      </button>
                                    ) : (
                                      '-'
                                    )}
                                  </td>
                                  <td className="p-3">
                                    <span
                                      className={`px-2 py-0.5 rounded-md text-[10px] font-bold ${
                                        l.status === 'Approved'
                                          ? 'bg-emerald-100 text-emerald-700'
                                          : l.status === 'Rejected'
                                          ? 'bg-rose-100 text-rose-700'
                                          : 'bg-amber-100 text-amber-700'
                                      }`}
                                    >
                                      {l.status}
                                    </span>
                                  </td>
                                  <td className="p-3 text-right space-x-1">
                                    {l.status === 'Pending' ? (
                                      <>
                                        <button
                                          onClick={async () => {
                                            await supabaseClient.from('lembur_presensi').update({ status: 'Approved' }).eq('id', l.id);
                                            showToast('Pengajuan lembur disetujui!', 'info');
                                            loadAdminLembur();
                                          }}
                                          className="px-2.5 py-1 rounded-lg bg-emerald-600 text-white font-bold text-xs cursor-pointer"
                                        >
                                          <i className="fa-solid fa-check"></i>
                                        </button>
                                        <button
                                          onClick={async () => {
                                            await supabaseClient.from('lembur_presensi').update({ status: 'Rejected' }).eq('id', l.id);
                                            showToast('Pengajuan lembur ditolak.', 'info');
                                            loadAdminLembur();
                                          }}
                                          className="px-2.5 py-1 rounded-lg bg-rose-600 text-white font-bold text-xs cursor-pointer"
                                        >
                                          <i className="fa-solid fa-xmark"></i>
                                        </button>
                                      </>
                                    ) : (
                                      <span className="text-slate-400 text-xs">Selesai</span>
                                    )}
                                  </td>
                                </tr>
                              ))
                            )}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  )}

                  {/* 7. PANEL JATAH LIBUR & CUTI */}
                  {activeAdminTab === 'jatah' && (
                    <div className="space-y-4">
                      <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                        <div>
                          <h3 className="font-extrabold text-base text-slate-900">Jatah Libur & Cuti Karyawan</h3>
                          <p className="text-xs text-slate-500">Monitoring dan penyesuaian kuota libur / cuti</p>
                        </div>
                        <button
                          onClick={async () => {
                            await supabaseClient.from('karyawan_presensi').update({ sisa_libur: 4 }).neq('id', '00000000-0000-0000-0000-000000000000');
                            showToast('Tutup buku berhasil! Seluruh jatah libur karyawan di-reset ke 4 hari.', 'success');
                            loadAdminKaryawan();
                          }}
                          className="px-3.5 py-2 bg-brand-600 hover:bg-brand-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-md cursor-pointer"
                        >
                          <i className="fa-solid fa-rotate-left"></i> Tutup Buku (Reset Libur 4 Hari)
                        </button>
                      </div>

                      <div className="overflow-x-auto border border-slate-200 rounded-2xl">
                        <table className="w-full text-xs text-left text-slate-600">
                          <thead className="bg-slate-50 text-slate-700 uppercase text-[10px] font-black border-b border-slate-200">
                            <tr>
                              <th className="p-3">Karyawan</th>
                              <th className="p-3">Divisi</th>
                              <th className="p-3">Sisa Libur</th>
                              <th className="p-3">Sisa Cuti</th>
                              <th className="p-3 text-right">Penyesuaian Manual</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-100 font-medium">
                            {adminKaryawanList.length === 0 ? (
                              <tr>
                                <td colSpan={5} className="p-4 text-center text-slate-400">Belum ada karyawan.</td>
                              </tr>
                            ) : (
                              adminKaryawanList.map((k, index) => (
                                <tr key={k.id} className="hover:bg-slate-50">
                                  <td className="p-3">
                                    <div className="flex items-center gap-2">
                                      <span className="inline-flex items-center justify-center min-w-5 h-5 px-1.5 rounded-md bg-slate-100 text-slate-700 text-[10px] font-extrabold border border-slate-200 shrink-0">
                                        {index + 1}
                                      </span>
                                      <span className="font-bold text-slate-900">{k.nama_lengkap}</span>
                                    </div>
                                  </td>
                                  <td className="p-3 text-slate-500">{k.divisi || 'Umum'}</td>
                                  <td className="p-3 font-extrabold text-sky-700">{k.sisa_libur ?? 4} Hari</td>
                                  <td className="p-3 font-extrabold text-purple-700">{k.sisa_cuti ?? 12} Hari</td>
                                  <td className="p-3 text-right space-x-1">
                                    <button
                                      onClick={async () => {
                                        const cur = k.sisa_libur ?? 4;
                                        await supabaseClient.from('karyawan_presensi').update({ sisa_libur: cur + 1 }).eq('id', k.id);
                                        loadAdminKaryawan();
                                      }}
                                      className="px-2 py-1 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold cursor-pointer"
                                    >
                                      + Libur
                                    </button>
                                    <button
                                      onClick={async () => {
                                        const cur = k.sisa_libur ?? 4;
                                        await supabaseClient.from('karyawan_presensi').update({ sisa_libur: Math.max(0, cur - 1) }).eq('id', k.id);
                                        loadAdminKaryawan();
                                      }}
                                      className="px-2 py-1 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold cursor-pointer"
                                    >
                                      - Libur
                                    </button>
                                    <button
                                      onClick={async () => {
                                        const cur = k.sisa_cuti ?? 12;
                                        await supabaseClient.from('karyawan_presensi').update({ sisa_cuti: cur + 1 }).eq('id', k.id);
                                        loadAdminKaryawan();
                                      }}
                                      className="px-2 py-1 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold cursor-pointer"
                                    >
                                      + Cuti
                                    </button>
                                    <button
                                      onClick={async () => {
                                        const cur = k.sisa_cuti ?? 12;
                                        await supabaseClient.from('karyawan_presensi').update({ sisa_cuti: Math.max(0, cur - 1) }).eq('id', k.id);
                                        loadAdminKaryawan();
                                      }}
                                      className="px-2 py-1 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold cursor-pointer"
                                    >
                                      - Cuti
                                    </button>
                                  </td>
                                </tr>
                              ))
                            )}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  )}

                  {/* 8. PANEL SIARAN PENGUMUMAN */}
                  {activeAdminTab === 'pengumuman' && (
                    <div className="space-y-4">
                      <div className="border-b border-slate-100 pb-3">
                        <h3 className="font-extrabold text-base text-slate-900">Siaran Pengumuman Realtime</h3>
                        <p className="text-xs text-slate-500">Kirim pengumuman langsung ke seluruh layar ponsel karyawan</p>
                      </div>

                      <form
                        onSubmit={async e => {
                          e.preventDefault();
                          const target = e.currentTarget;
                          const judul = (target.elements.namedItem('annJudul') as HTMLInputElement).value.trim();
                          const isi = (target.elements.namedItem('annIsi') as HTMLTextAreaElement).value.trim();
                          await supabaseClient.from('pengumuman_presensi').insert([{
                            judul,
                            isi,
                            author: currentAdmin.nama_lengkap || 'Admin'
                          }]);
                          showToast('Pengumuman berhasil disiarkan!', 'success');
                          target.reset();
                          loadAdminPengumuman();
                        }}
                        className="p-4 bg-slate-50 rounded-2xl border border-slate-200 space-y-3"
                      >
                        <div>
                          <label className="block text-xs font-bold text-slate-700 mb-1">Judul Pengumuman</label>
                          <input
                            name="annJudul"
                            type="text"
                            required
                            placeholder="Contoh: Jadwal Rapat Akbar Hari Ini"
                            className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs font-medium"
                          />
                        </div>
                        <div>
                          <label className="block text-xs font-bold text-slate-700 mb-1">Isi Pengumuman</label>
                          <textarea
                            name="annIsi"
                            required
                            rows={3}
                            placeholder="Tulis rincian pengumuman..."
                            className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs"
                          ></textarea>
                        </div>
                        <button
                          type="submit"
                          className="px-4 py-2 bg-brand-600 hover:bg-brand-700 text-white rounded-xl text-xs font-bold shadow-md flex items-center gap-1.5 cursor-pointer"
                        >
                          <i className="fa-solid fa-bullhorn"></i> Siarkan Pengumuman Sekarang
                        </button>
                      </form>

                      <div className="space-y-3 pt-2">
                        {adminPengumumanList.length === 0 ? (
                          <div className="text-xs text-slate-400 text-center py-4">Belum ada siaran pengumuman.</div>
                        ) : (
                          adminPengumumanList.map(p => (
                            <div key={p.id} className="p-4 bg-slate-50 rounded-2xl border border-slate-200 flex items-start justify-between gap-3">
                              <div>
                                <h5 className="font-bold text-sm text-slate-900">{p.judul}</h5>
                                <p className="text-xs text-slate-600 mt-1">{p.isi}</p>
                                <span className="text-[10px] text-slate-400 mt-2 block">
                                  {new Date(p.created_at).toLocaleString('id-ID')}
                                </span>
                              </div>
                              <button
                                onClick={() => {
                                  setDeleteConfirmMsg('Hapus pengumuman ini?');
                                  setDeleteActionFn(() => async () => {
                                    await supabaseClient.from('pengumuman_presensi').delete().eq('id', p.id);
                                    showToast('Pengumuman berhasil dihapus!', 'success');
                                    loadAdminPengumuman();
                                  });
                                  setModalConfirmDeleteOpen(true);
                                }}
                                className="text-rose-500 hover:text-rose-700 p-1 text-xs cursor-pointer"
                              >
                                <i className="fa-solid fa-trash"></i>
                              </button>
                            </div>
                          ))
                        )}
                      </div>
                    </div>
                  )}

                  {/* 9. PANEL MASTER ADMIN */}
                  {activeAdminTab === 'useradmin' && (
                    <div className="space-y-4">
                      <div className="flex items-center justify-between border-b border-slate-100 pb-3">
                        <div>
                          <h3 className="font-extrabold text-base text-slate-900">Master Akun Admin</h3>
                          <p className="text-xs text-slate-500">Kelola pengguna dengan hak akses dashboard admin</p>
                        </div>
                        <button
                          onClick={() => {
                            setEditingAdminId(null);
                            setAdminUserNama('');
                            setAdminUserUsername('');
                            setAdminUserEmail('');
                            setAdminUserPass('');
                            setModalAdminOpen(true);
                          }}
                          className="px-3.5 py-2 bg-brand-600 hover:bg-brand-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-md cursor-pointer"
                        >
                          <i className="fa-solid fa-user-plus"></i> Tambah Admin
                        </button>
                      </div>

                      <div className="overflow-x-auto border border-slate-200 rounded-2xl">
                        <table className="w-full text-xs text-left text-slate-600">
                          <thead className="bg-slate-50 text-slate-700 uppercase text-[10px] font-black border-b border-slate-200">
                            <tr>
                              <th className="p-3">Nama Lengkap</th>
                              <th className="p-3">Username</th>
                              <th className="p-3">Email</th>
                              <th className="p-3">Role</th>
                              <th className="p-3 text-right">Aksi</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-100 font-medium">
                            {adminUsersList.length === 0 ? (
                              <tr>
                                <td colSpan={5} className="p-4 text-center text-slate-400">Belum ada akun admin.</td>
                              </tr>
                            ) : (
                              adminUsersList.map(a => (
                                <tr key={a.id} className="hover:bg-slate-50">
                                  <td className="p-3 font-bold text-slate-900">{a.nama_lengkap}</td>
                                  <td className="p-3 font-mono text-brand-600">{a.username}</td>
                                  <td className="p-3 text-slate-500">{a.email}</td>
                                  <td className="p-3">
                                    <span className="px-2 py-0.5 rounded-full text-[10px] font-bold bg-indigo-100 text-indigo-700 uppercase">
                                      {a.role || 'Admin'}
                                    </span>
                                  </td>
                                  <td className="p-3 text-right space-x-1">
                                    <button
                                      onClick={() => {
                                        setEditingAdminId(a.id);
                                        setAdminUserNama(a.nama_lengkap);
                                        setAdminUserUsername(a.username);
                                        setAdminUserEmail(a.email);
                                        setAdminUserPass('');
                                        setModalAdminOpen(true);
                                      }}
                                      className="px-2 py-1 rounded-lg bg-indigo-50 text-indigo-600 font-bold cursor-pointer"
                                    >
                                      <i className="fa-solid fa-pen"></i>
                                    </button>
                                    <button
                                      onClick={() => {
                                        setDeleteConfirmMsg(`Hapus admin ${a.username}?`);
                                        setDeleteActionFn(() => async () => {
                                          await supabaseClient.from('admin_presensi').delete().eq('id', a.id);
                                          showToast('Admin berhasil dihapus!', 'success');
                                          loadAdminUsers();
                                        });
                                        setModalConfirmDeleteOpen(true);
                                      }}
                                      className="px-2 py-1 rounded-lg bg-rose-50 text-rose-600 font-bold cursor-pointer"
                                    >
                                      <i className="fa-solid fa-trash"></i>
                                    </button>
                                  </td>
                                </tr>
                              ))
                            )}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  )}
                </div>
              </div>
            )}
          </section>
        )}
      </main>

      {/* ---------------- MODALS SECTION ---------------- */}

      {/* MODAL 1: PRESENSI MULTI-STEP WIZARD */}
      {modalPresensiOpen && (
        <div className="fixed inset-0 z-50 bg-black/70 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-white max-w-md w-full rounded-3xl overflow-hidden shadow-2xl flex flex-col max-h-[92vh]">
            {/* Header with Step indicator */}
            <div className="px-5 py-3.5 bg-slate-900 text-white flex items-center justify-between border-b border-slate-800">
              <div className="flex items-center gap-2.5">
                <div className="w-7 h-7 rounded-xl bg-brand-500/20 text-brand-400 flex items-center justify-center text-xs font-black">
                  <span>{presensiStep}</span>
                </div>
                <div>
                  <h3 className="font-extrabold text-xs">
                    {presensiMode === 'masuk' ? 'Presensi Masuk' : 'Presensi Keluar'}
                  </h3>
                  <span className="text-[10px] text-slate-400 font-medium">
                    {presensiStep === 1
                      ? 'Langkah 1: Pilih Lokasi'
                      : presensiStep === 2
                      ? presensiMode === 'masuk'
                        ? 'Langkah 2: Pilih Shift'
                        : 'Langkah 2: Pilihan Status Pulang'
                      : 'Langkah 3: Foto Selfie & Kirim'}
                  </span>
                </div>
              </div>

              <div className="flex items-center gap-3">
                <div className="flex items-center gap-1.5">
                  <div className={`h-1.5 rounded-full transition-all duration-300 ${presensiStep === 1 ? 'w-6 bg-brand-600' : 'w-2 bg-emerald-500'}`}></div>
                  <div className={`h-1.5 rounded-full transition-all duration-300 ${presensiStep === 2 ? 'w-6 bg-brand-600' : presensiStep > 2 ? 'w-2 bg-emerald-500' : 'w-2 bg-slate-700'}`}></div>
                  <div className={`h-1.5 rounded-full transition-all duration-300 ${presensiStep === 3 ? 'w-6 bg-brand-600' : 'w-2 bg-slate-700'}`}></div>
                </div>
                <button onClick={closeModalPresensi} className="text-slate-400 hover:text-white p-1 transition-all cursor-pointer">
                  <i className="fa-solid fa-xmark text-lg"></i>
                </button>
              </div>
            </div>

            <div className="p-5 overflow-y-auto flex-1 space-y-4">
              {/* SLIDE 1: Lokasi & Map */}
              {presensiStep === 1 && (
                <div className="space-y-4">
                  <div>
                    <label className="block text-xs font-bold text-slate-700 uppercase mb-1.5">Pilih Lokasi Kantor Penugasan</label>
                    <select
                      value={selectedOfficeId}
                      onChange={e => setSelectedOfficeId(e.target.value)}
                      className="w-full px-3.5 py-3 bg-slate-50 border border-slate-200 rounded-2xl text-xs font-bold focus:outline-none focus:ring-2 focus:ring-brand-500 transition-all shadow-sm"
                    >
                      {cachedLokasi.map(l => (
                        <option key={l.id} value={l.id}>
                          {l.nama_lokasi} (Radius {l.radius_meter}m)
                        </option>
                      ))}
                    </select>
                  </div>

                  <div>
                    <div className="flex items-center justify-between mb-1.5">
                      <span className="text-xs font-bold text-slate-700 uppercase">Verifikasi Geofence GPS</span>
                      <span className={gpsBadgeClass}>{gpsStatusText}</span>
                    </div>
                    <div ref={mapContainerRef} className="w-full h-44 rounded-2xl border border-slate-200 overflow-hidden shadow-inner"></div>
                    <p className="text-[10px] text-slate-400 mt-1 italic">
                      *Pastikan posisi GPS Anda berada di dalam lingkaran radius kantor yang ditentukan.
                    </p>
                  </div>

                  <div className="pt-2 flex items-center justify-between gap-2 border-t border-slate-100">
                    <button type="button" onClick={closeModalPresensi} className="px-4 py-2.5 rounded-xl text-xs font-bold text-slate-600 hover:bg-slate-100 transition-all cursor-pointer">
                      Batal
                    </button>
                    <button
                      type="button"
                      onClick={() => setPresensiStep(2)}
                      className="px-5 py-2.5 rounded-xl text-xs font-bold bg-brand-600 hover:bg-brand-700 text-white shadow-md shadow-brand-500/25 transition-all flex items-center gap-1.5 cursor-pointer"
                    >
                      <span>{presensiMode === 'masuk' ? 'Lanjut: Pilih Shift' : 'Lanjut: Pilihan Status'}</span>
                      <i className="fa-solid fa-arrow-right text-[11px]"></i>
                    </button>
                  </div>
                </div>
              )}

              {/* SLIDE 2 (MASUK): Shift Kerja */}
              {presensiStep === 2 && presensiMode === 'masuk' && (
                <div className="space-y-4">
                  <div className="p-4 bg-brand-50/60 rounded-2xl border border-brand-100">
                    <div className="text-xs font-bold text-brand-900 mb-1 flex items-center gap-1.5">
                      <i className="fa-solid fa-business-time text-brand-600"></i>
                      <span>Pilih Jadwal Shift Tugas Hari Ini</span>
                    </div>
                    <p className="text-[11px] text-brand-700/80 leading-relaxed">
                      Shift menentukan perhitungan toleransi keterlambatan jam masuk Anda.
                    </p>
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-700 uppercase mb-1.5">Shift Kerja Anda</label>
                    <select
                      value={selectedShiftName}
                      onChange={e => setSelectedShiftName(e.target.value)}
                      className="w-full px-3.5 py-3 bg-slate-50 border border-slate-200 rounded-2xl text-xs font-bold focus:outline-none focus:ring-2 focus:ring-brand-500 transition-all shadow-sm"
                    >
                      {currentUser?.shifts && currentUser.shifts.length > 0 ? (
                        currentUser.shifts.map(s => <option key={s} value={s}>{s}</option>)
                      ) : (
                        <option value="Reguler">Reguler</option>
                      )}
                    </select>
                  </div>

                  <div className="pt-4 flex items-center justify-between gap-2 border-t border-slate-100">
                    <button type="button" onClick={() => setPresensiStep(1)} className="px-4 py-2.5 rounded-xl text-xs font-bold text-slate-600 hover:bg-slate-100 transition-all flex items-center gap-1.5 cursor-pointer">
                      <i className="fa-solid fa-arrow-left text-[11px]"></i>
                      <span>Kembali</span>
                    </button>
                    <button
                      type="button"
                      onClick={() => setPresensiStep(3)}
                      className="px-5 py-2.5 rounded-xl text-xs font-bold bg-brand-600 hover:bg-brand-700 text-white shadow-md shadow-brand-500/25 transition-all flex items-center gap-1.5 cursor-pointer"
                    >
                      <span>Lanjut: Foto Selfie</span>
                      <i className="fa-solid fa-arrow-right text-[11px]"></i>
                    </button>
                  </div>
                </div>
              )}

              {/* SLIDE 2 (KELUAR): Pilihan Status Pulang */}
              {presensiStep === 2 && presensiMode === 'keluar' && (
                <div className="space-y-4">
                  <div className="p-4 bg-rose-50/60 rounded-2xl border border-rose-100">
                    <div className="text-xs font-bold text-rose-900 mb-1 flex items-center gap-1.5">
                      <i className="fa-solid fa-arrow-right-from-bracket text-rose-600"></i>
                      <span>Status Kepulangan / Selesai Tugas</span>
                    </div>
                    <p className="text-[11px] text-rose-700/80 leading-relaxed">
                      Pilih alasan kepulangan Anda: PULANG, IZIN, atau SAKIT.
                    </p>
                  </div>

                  <div>
                    <label className="block text-xs font-bold text-slate-700 uppercase mb-1.5">Pilihan Status Pulang</label>
                    <select
                      value={statusKeluar}
                      onChange={e => setStatusKeluar(e.target.value as any)}
                      className="w-full px-3.5 py-3 bg-slate-50 border border-slate-200 rounded-2xl text-xs font-bold focus:outline-none focus:ring-2 focus:ring-brand-500 shadow-sm"
                    >
                      <option value="PULANG">PULANG (Selesai Jam Kerja)</option>
                      <option value="IZIN">IZIN (Meninggalkan Kantor / Urusan Luar)</option>
                      <option value="SAKIT">SAKIT (Izin Pulang Karena Sakit)</option>
                      <option value="LIBUR">LIBUR (Jadwal Libur)</option>
                      <option value="CUTI">CUTI (Pengajuan Cuti)</option>
                    </select>
                  </div>

                  {statusKeluar !== 'PULANG' && (
                    <div>
                      <label className="block text-xs font-bold text-slate-700 uppercase mb-1.5">
                        Keterangan / Alasan Lengkap <span className="text-rose-500">*</span>
                      </label>
                      <textarea
                        value={presensiKeterangan}
                        onChange={e => setPresensiKeterangan(e.target.value)}
                        rows={3}
                        placeholder="Tuliskan keterangan lengkap di sini..."
                        className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-2xl text-xs font-medium focus:outline-none focus:ring-2 focus:ring-brand-500"
                      ></textarea>
                    </div>
                  )}

                  <div className="pt-4 flex items-center justify-between gap-2 border-t border-slate-100">
                    <button type="button" onClick={() => setPresensiStep(1)} className="px-4 py-2.5 rounded-xl text-xs font-bold text-slate-600 hover:bg-slate-100 transition-all flex items-center gap-1.5 cursor-pointer">
                      <i className="fa-solid fa-arrow-left text-[11px]"></i>
                      <span>Kembali</span>
                    </button>
                    <button
                      type="button"
                      onClick={() => {
                        if ((statusKeluar === 'IZIN' || statusKeluar === 'SAKIT') && !presensiKeterangan.trim()) {
                          showToast(`Harap tuliskan keterangan / alasan untuk status ${statusKeluar}`, 'error');
                          return;
                        }
                        setPresensiStep(3);
                      }}
                      className="px-5 py-2.5 rounded-xl text-xs font-bold bg-brand-600 hover:bg-brand-700 text-white shadow-md shadow-brand-500/25 transition-all flex items-center gap-1.5 cursor-pointer"
                    >
                      <span>Lanjut: Foto Selfie</span>
                      <i className="fa-solid fa-arrow-right text-[11px]"></i>
                    </button>
                  </div>
                </div>
              )}

              {/* SLIDE 3: Camera & Kirim */}
              {presensiStep === 3 && (
                <div className="space-y-4">
                  <div className="relative w-full aspect-square rounded-2xl overflow-hidden bg-black flex items-center justify-center border-2 border-slate-800 shadow-inner">
                    <video
                      ref={videoRef}
                      autoPlay
                      playsInline
                      className={`w-full h-full object-cover ${capturedPhoto ? 'hidden' : 'block'}`}
                    ></video>
                    <canvas ref={canvasRef} className="hidden"></canvas>
                    {capturedPhoto && (
                      <img src={capturedPhoto} alt="Captured preview" className="w-full h-full object-cover" />
                    )}

                    <div className="absolute bottom-3 inset-x-0 flex items-center justify-center gap-3 z-10">
                      {!capturedPhoto ? (
                        <button
                          type="button"
                          onClick={takeSnapshot}
                          className="w-14 h-14 rounded-full bg-white text-slate-900 flex items-center justify-center shadow-lg active:scale-95 transition-all cursor-pointer"
                        >
                          <i className="fa-solid fa-camera text-xl text-brand-600"></i>
                        </button>
                      ) : (
                        <button
                          type="button"
                          onClick={retakeSnapshot}
                          className="px-4 py-2 rounded-full bg-black/60 backdrop-blur-md text-white text-xs font-bold border border-white/20 active:scale-95 transition-all cursor-pointer"
                        >
                          <i className="fa-solid fa-rotate-right mr-1"></i> Foto Ulang
                        </button>
                      )}
                    </div>
                  </div>

                  <div className="text-center p-2 rounded-xl text-xs font-bold bg-slate-100 text-slate-600">
                    {gpsStatusText}
                  </div>

                  <div className="pt-2 flex items-center justify-between gap-2 border-t border-slate-100">
                    <button type="button" onClick={() => setPresensiStep(2)} className="px-4 py-2.5 rounded-xl text-xs font-bold text-slate-600 hover:bg-slate-100 transition-all flex items-center gap-1.5 cursor-pointer">
                      <i className="fa-solid fa-arrow-left text-[11px]"></i>
                      <span>Kembali</span>
                    </button>
                    <button
                      type="button"
                      disabled={!isInsideGeofence || !capturedPhoto || isPresensiSubmitting}
                      onClick={submitPresensi}
                      className="px-5 py-2.5 rounded-xl text-xs font-bold bg-brand-600 hover:bg-brand-700 text-white shadow-md shadow-brand-500/25 transition-all flex items-center gap-1.5 disabled:opacity-50 disabled:pointer-events-none cursor-pointer"
                    >
                      {isPresensiSubmitting ? (
                        <>
                          <i className="fa-solid fa-spinner fa-spin"></i>
                          <span>Menyimpan...</span>
                        </>
                      ) : (
                        <>
                          <span>Kirim Presensi</span>
                          <i className="fa-solid fa-paper-plane text-[11px]"></i>
                        </>
                      )}
                    </button>
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* MODAL 2: PENGAJUAN LEMBUR */}
      {modalLemburOpen && (
        <div className="fixed inset-0 z-50 bg-black/70 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-white max-w-sm w-full rounded-3xl p-5 shadow-2xl border border-slate-100 space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <div className="flex items-center gap-2">
                <div className="w-8 h-8 rounded-xl bg-orange-50 text-orange-600 flex items-center justify-center text-sm">
                  <i className="fa-solid fa-clock-rotate-left"></i>
                </div>
                <div>
                  <h3 className="font-extrabold text-sm text-slate-900">Pengajuan Lembur</h3>
                  <p className="text-[10px] text-slate-400 font-medium">Unggah bukti foto, durasi & keterangan</p>
                </div>
              </div>
              <button onClick={() => setModalLemburOpen(false)} className="text-slate-400 hover:text-slate-700 text-sm p-1 cursor-pointer">
                <i className="fa-solid fa-xmark"></i>
              </button>
            </div>

            <form onSubmit={submitLembur} className="space-y-3.5">
              <div>
                <label className="block text-xs font-bold text-slate-700 uppercase mb-1.5">Foto Bukti Lembur (Opsional)</label>
                <input
                  type="file"
                  id="lemburFotoInputEl"
                  accept="image/*"
                  onChange={e => {
                    const file = e.target.files?.[0];
                    if (!file) return;
                    const r = new FileReader();
                    r.onload = ev => setLemburFoto(ev.target?.result as string);
                    r.readAsDataURL(file);
                  }}
                  className="hidden"
                />

                {!lemburFoto ? (
                  <div
                    onClick={() => document.getElementById('lemburFotoInputEl')?.click()}
                    className="border-2 border-dashed border-slate-300 hover:border-orange-500 rounded-2xl p-5 flex flex-col items-center justify-center cursor-pointer transition-all bg-slate-50 hover:bg-orange-50/20 text-center"
                  >
                    <div className="w-12 h-12 rounded-2xl bg-white shadow-sm border border-slate-200 flex items-center justify-center text-orange-600 mb-2">
                      <i className="fa-solid fa-camera text-xl"></i>
                    </div>
                    <span className="text-xs font-bold text-slate-700">Pilih Foto</span>
                    <span className="text-[10px] text-slate-400 mt-0.5">Bukti pekerjaan lembur (JPG, PNG)</span>
                  </div>
                ) : (
                  <div className="relative rounded-2xl overflow-hidden border border-slate-200 bg-slate-900 aspect-video flex items-center justify-center">
                    <img src={lemburFoto} alt="Preview lembur" className="max-h-full max-w-full object-contain" />
                    <button
                      type="button"
                      onClick={() => setLemburFoto(null)}
                      className="absolute top-2 right-2 px-3 py-1 rounded-xl bg-black/70 hover:bg-rose-600 text-white text-[11px] font-bold border border-white/20 transition-all flex items-center gap-1 cursor-pointer"
                    >
                      <i className="fa-solid fa-rotate-left"></i> Ganti Foto
                    </button>
                  </div>
                )}
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-700 uppercase mb-1.5">
                  Durasi Lembur (Jam) <span className="text-rose-500">*</span>
                </label>
                <input
                  type="number"
                  min="1"
                  max="12"
                  value={lemburDurasi}
                  onChange={e => setLemburDurasi(e.target.value)}
                  placeholder="Contoh: 2"
                  required
                  className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-2xl text-xs font-bold focus:outline-none focus:ring-2 focus:ring-orange-500 transition-all"
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-700 uppercase mb-1.5">
                  Keterangan Pekerjaan <span className="text-rose-500">*</span>
                </label>
                <textarea
                  value={lemburKeterangan}
                  onChange={e => setLemburKeterangan(e.target.value)}
                  required
                  rows={3}
                  placeholder="Uraikan pekerjaan lembur yang dilakukan..."
                  className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-2xl text-xs font-medium focus:outline-none focus:ring-2 focus:ring-orange-500 transition-all"
                ></textarea>
              </div>

              <div className="pt-2 flex items-center justify-end gap-2 border-t border-slate-100">
                <button type="button" onClick={() => setModalLemburOpen(false)} className="px-4 py-2.5 rounded-xl text-xs font-bold text-slate-600 hover:bg-slate-100 transition-all cursor-pointer">
                  Batal
                </button>
                <button
                  type="submit"
                  disabled={lemburSubmitting}
                  className="px-5 py-2.5 rounded-xl text-xs font-bold bg-orange-600 hover:bg-orange-700 text-white shadow-md shadow-orange-500/25 transition-all flex items-center gap-1.5 cursor-pointer disabled:opacity-50"
                >
                  {lemburSubmitting ? (
                    <>
                      <i className="fa-solid fa-spinner fa-spin"></i>
                      <span>Mengirim...</span>
                    </>
                  ) : (
                    <>
                      <span>Kirim Pengajuan</span>
                      <i className="fa-solid fa-paper-plane text-[11px]"></i>
                    </>
                  )}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* MODAL 3: PENGAJUAN SAKIT / IZIN / CUTI */}
      {modalIzinOpen && (
        <div className="fixed inset-0 z-50 bg-black/70 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-white max-w-sm w-full rounded-3xl p-5 shadow-2xl border border-slate-100 space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <div className="flex items-center gap-2">
                <div
                  className={`w-8 h-8 rounded-xl flex items-center justify-center text-sm ${
                    izinType === 'SAKIT'
                      ? 'bg-amber-50 text-amber-600'
                      : izinType === 'IZIN'
                      ? 'bg-amber-50 text-amber-600'
                      : 'bg-purple-50 text-purple-600'
                  }`}
                >
                  <i
                    className={`fa-solid ${
                      izinType === 'SAKIT'
                        ? 'fa-briefcase-medical'
                        : izinType === 'IZIN'
                        ? 'fa-envelope-open-text'
                        : 'fa-calendar-check'
                    }`}
                  ></i>
                </div>
                <div>
                  <h3 className="font-extrabold text-sm text-slate-900">
                    Pengajuan {izinType === 'SAKIT' ? 'Sakit' : izinType === 'IZIN' ? 'Izin' : 'Cuti'}
                  </h3>
                  <p className="text-[10px] text-slate-400 font-medium">Unggah foto bukti / surat keterangan</p>
                </div>
              </div>
              <button onClick={() => setModalIzinOpen(false)} className="text-slate-400 hover:text-slate-700 text-sm p-1 cursor-pointer">
                <i className="fa-solid fa-xmark"></i>
              </button>
            </div>

            <form onSubmit={submitIzin} className="space-y-3.5">
              <div>
                <label className="block text-xs font-bold text-slate-700 uppercase mb-1.5">Unggah Foto Bukti / Surat</label>
                <input
                  type="file"
                  id="izinFotoInputEl"
                  accept="image/*"
                  onChange={e => {
                    const file = e.target.files?.[0];
                    if (!file) return;
                    const r = new FileReader();
                    r.onload = ev => setIzinFoto(ev.target?.result as string);
                    r.readAsDataURL(file);
                  }}
                  className="hidden"
                />

                {!izinFoto ? (
                  <div
                    onClick={() => document.getElementById('izinFotoInputEl')?.click()}
                    className="border-2 border-dashed border-slate-300 hover:border-brand-500 rounded-2xl p-5 flex flex-col items-center justify-center cursor-pointer transition-all bg-slate-50 hover:bg-brand-50/20 text-center"
                  >
                    <div className="w-12 h-12 rounded-2xl bg-white shadow-sm border border-slate-200 flex items-center justify-center text-brand-600 mb-2">
                      <i className="fa-solid fa-camera text-xl"></i>
                    </div>
                    <span className="text-xs font-bold text-slate-700">Pilih Foto</span>
                    <span className="text-[10px] text-slate-400 mt-0.5">Surat dokter / bukti izin (JPG, PNG)</span>
                  </div>
                ) : (
                  <div className="relative rounded-2xl overflow-hidden border border-slate-200 bg-slate-900 aspect-video flex items-center justify-center">
                    <img src={izinFoto} alt="Preview bukti" className="max-h-full max-w-full object-contain" />
                    <button
                      type="button"
                      onClick={() => setIzinFoto(null)}
                      className="absolute top-2 right-2 px-3 py-1 rounded-xl bg-black/70 hover:bg-rose-600 text-white text-[11px] font-bold border border-white/20 transition-all flex items-center gap-1 cursor-pointer"
                    >
                      <i className="fa-solid fa-rotate-left"></i> Ganti Foto
                    </button>
                  </div>
                )}
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-700 uppercase mb-1.5">
                  Keterangan / Alasan <span className="text-rose-500">*</span>
                </label>
                <textarea
                  value={izinKeterangan}
                  onChange={e => setIzinKeterangan(e.target.value)}
                  required
                  rows={3}
                  placeholder="Tuliskan keterangan lengkap di sini..."
                  className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-2xl text-xs font-medium focus:outline-none focus:ring-2 focus:ring-brand-500 transition-all"
                ></textarea>
              </div>

              <div className="pt-2 flex items-center justify-end gap-2 border-t border-slate-100">
                <button type="button" onClick={() => setModalIzinOpen(false)} className="px-4 py-2.5 rounded-xl text-xs font-bold text-slate-600 hover:bg-slate-100 transition-all cursor-pointer">
                  Batal
                </button>
                <button
                  type="submit"
                  disabled={izinSubmitting}
                  className="px-5 py-2.5 rounded-xl text-xs font-bold bg-brand-600 hover:bg-brand-700 text-white shadow-md shadow-brand-500/25 transition-all flex items-center gap-1.5 cursor-pointer disabled:opacity-50"
                >
                  {izinSubmitting ? (
                    <>
                      <i className="fa-solid fa-spinner fa-spin"></i>
                      <span>Mengirim...</span>
                    </>
                  ) : (
                    <>
                      <span>Kirim Pengajuan</span>
                      <i className="fa-solid fa-paper-plane text-[11px]"></i>
                    </>
                  )}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* MODAL 4: LIBUR DIRECT CONFIRM */}
      {modalLiburOpen && (
        <div className="fixed inset-0 z-50 bg-black/70 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-white max-w-sm w-full rounded-3xl p-5 shadow-2xl border border-slate-100 space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <div className="flex items-center gap-2">
                <div className="w-8 h-8 rounded-xl bg-sky-50 text-sky-600 flex items-center justify-center text-sm">
                  <i className="fa-solid fa-umbrella-beach"></i>
                </div>
                <div>
                  <h3 className="font-extrabold text-sm text-slate-900">Konfirmasi Libur</h3>
                  <p className="text-[10px] text-slate-400 font-medium">Pengambilan jadwal libur rutin</p>
                </div>
              </div>
              <button onClick={() => setModalLiburOpen(false)} className="text-slate-400 hover:text-slate-700 text-sm p-1 cursor-pointer">
                <i className="fa-solid fa-xmark"></i>
              </button>
            </div>

            <form onSubmit={submitLiburDirect} className="space-y-3.5">
              <div className="p-3.5 bg-sky-50 border border-sky-100 rounded-2xl flex items-center justify-between text-xs">
                <span className="font-bold text-sky-900">Sisa Jatah Libur:</span>
                <span className="font-black text-sky-700 text-sm">{currentUser?.sisa_libur ?? 4} Hari</span>
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-700 uppercase mb-1.5">Catatan Tambahan (Opsional)</label>
                <textarea
                  value={liburKeterangan}
                  onChange={e => setLiburKeterangan(e.target.value)}
                  rows={2}
                  placeholder="Catatan libur (opsional)..."
                  className="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-200 rounded-2xl text-xs font-medium focus:outline-none focus:ring-2 focus:ring-brand-500 transition-all"
                ></textarea>
              </div>

              <div className="pt-2 flex items-center justify-end gap-2 border-t border-slate-100">
                <button type="button" onClick={() => setModalLiburOpen(false)} className="px-4 py-2.5 rounded-xl text-xs font-bold text-slate-600 hover:bg-slate-100 transition-all cursor-pointer">
                  Batal
                </button>
                <button
                  type="submit"
                  disabled={liburSubmitting}
                  className="px-5 py-2.5 rounded-xl text-xs font-bold bg-sky-600 hover:bg-sky-700 text-white shadow-md shadow-sky-500/25 transition-all flex items-center gap-1.5 cursor-pointer disabled:opacity-50"
                >
                  {liburSubmitting ? (
                    <>
                      <i className="fa-solid fa-spinner fa-spin"></i>
                      <span>Menyimpan...</span>
                    </>
                  ) : (
                    <>
                      <span>Konfirmasi Libur</span>
                      <i className="fa-solid fa-check text-[11px]"></i>
                    </>
                  )}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* MODAL 5: CUTI LOCKED */}
      {modalCutiLockedOpen && (
        <div className="fixed inset-0 z-50 bg-black/70 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-white max-w-sm w-full rounded-3xl p-6 shadow-2xl border border-slate-100 text-center space-y-4">
            <div className="w-14 h-14 rounded-2xl bg-amber-50 text-amber-600 flex items-center justify-center mx-auto shadow-inner text-2xl">
              <i className="fa-solid fa-lock"></i>
            </div>
            <div>
              <h3 className="font-extrabold text-base text-slate-900">Akses Cuti Terkunci</h3>
              <p className="text-xs text-slate-600 mt-2 leading-relaxed">
                🚧 Akses cuti untuk divisi kamu masih terkunci. Semoga Suryatokoku terus bertumbuh, semua tim makin sejahtera, dan fitur ini segera berubah dari angan-angan menjadi kenyataan. Semangat 🦾!
              </p>
            </div>
            <div className="pt-2">
              <button
                type="button"
                onClick={() => setModalCutiLockedOpen(false)}
                className="w-full py-2.5 px-4 rounded-xl text-xs font-bold text-white bg-brand-600 hover:bg-brand-700 shadow-md shadow-brand-500/25 transition-all cursor-pointer"
              >
                Siap, Tetap Semangat! 🙌
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL 6: PWA GUIDE */}
      {modalPwaGuideOpen && (
        <div className="fixed inset-0 z-50 bg-black/70 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-white max-w-sm w-full rounded-3xl p-5 shadow-2xl border border-slate-100 space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <div className="flex items-center gap-2">
                <div className="w-7 h-7 rounded-lg bg-brand-50 text-brand-600 flex items-center justify-center text-xs">
                  <i className="fa-solid fa-mobile-screen"></i>
                </div>
                <h3 className="font-extrabold text-sm text-slate-900">Pasang ke Layar Utama</h3>
              </div>
              <button onClick={() => setModalPwaGuideOpen(false)} className="text-slate-400 hover:text-slate-700 text-sm p-1 cursor-pointer">
                <i className="fa-solid fa-xmark"></i>
              </button>
            </div>

            <div className="space-y-3 text-xs text-slate-600">
              <p className="leading-relaxed">
                Agar aplikasi berjalan <b>layar penuh (fullscreen)</b> seperti aplikasi bawaan HP:
              </p>

              <div className="p-3 bg-slate-50 rounded-2xl border border-slate-100 space-y-2">
                <div className="font-bold text-slate-800 flex items-center gap-1.5 text-[11px] uppercase tracking-wider">
                  <i className="fa-brands fa-apple text-slate-900"></i> Untuk Pengguna iPhone / Safari:
                </div>
                <ol className="list-decimal list-inside space-y-1 text-[11px] text-slate-600 pl-1">
                  <li>
                    Tekan tombol <b>Bagikan / Share</b> (<i className="fa-solid fa-arrow-up-from-bracket text-brand-600"></i>) di bawah layar Safari.
                  </li>
                  <li>
                    Gulir ke bawah, pilih <b>"Tambahkan ke Layar Utama"</b> (<i>Add to Home Screen</i>).
                  </li>
                  <li>
                    Tekan <b>Tambah</b> di pojok kanan atas.
                  </li>
                </ol>
              </div>

              <div className="p-3 bg-slate-50 rounded-2xl border border-slate-100 space-y-2">
                <div className="font-bold text-slate-800 flex items-center gap-1.5 text-[11px] uppercase tracking-wider">
                  <i className="fa-brands fa-android text-emerald-600"></i> Untuk Android / Chrome:
                </div>
                <ol className="list-decimal list-inside space-y-1 text-[11px] text-slate-600 pl-1">
                  <li>
                    Tekan tombol titik tiga (<i className="fa-solid fa-ellipsis-vertical text-slate-700"></i>) di kanan atas browser.
                  </li>
                  <li>
                    Pilih <b>"Instal aplikasi"</b> atau <b>"Tambahkan ke Layar Utama"</b>.
                  </li>
                  <li>Konfirmasi pemasangan shortcut.</li>
                </ol>
              </div>
            </div>

            <button
              type="button"
              onClick={() => setModalPwaGuideOpen(false)}
              className="w-full py-2.5 bg-slate-900 hover:bg-black text-white rounded-xl text-xs font-bold transition-all cursor-pointer"
            >
              Saya Mengerti
            </button>
          </div>
        </div>
      )}

      {/* MODAL 7: PHOTO PREVIEW */}
      {photoPreviewModalOpen && (
        <div
          onClick={e => {
            if (e.target === e.currentTarget) setPhotoPreviewModalOpen(false);
          }}
          className="fixed inset-0 z-50 bg-black/80 backdrop-blur-sm flex items-center justify-center p-4"
        >
          <div className="bg-white max-w-lg w-full rounded-3xl overflow-hidden shadow-2xl flex flex-col max-h-[90vh]">
            <div className="px-5 py-3.5 bg-slate-900 text-white flex items-center justify-between">
              <div className="min-w-0 pr-3">
                <h3 className="font-extrabold text-sm truncate">{previewPhotoTitle}</h3>
                <p className="text-[11px] text-slate-400 font-medium truncate mt-0.5">{previewPhotoSubtitle}</p>
              </div>
              <div className="flex items-center gap-1.5 shrink-0">
                <a
                  href={previewPhotoUrl}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="px-2.5 py-1.5 rounded-xl bg-white/10 hover:bg-white/20 text-white flex items-center gap-1.5 text-xs font-bold transition-all"
                  title="Buka foto di tab baru"
                >
                  <i className="fa-solid fa-arrow-up-right-from-square text-[11px]"></i>
                  <span className="text-[10px]">Tab Baru</span>
                </a>
                <button
                  type="button"
                  onClick={() => setPhotoPreviewModalOpen(false)}
                  className="w-8 h-8 rounded-xl bg-white/10 hover:bg-rose-600 text-white flex items-center justify-center text-sm transition-all cursor-pointer"
                  title="Tutup"
                >
                  <i className="fa-solid fa-xmark"></i>
                </button>
              </div>
            </div>

            <div className="relative bg-slate-950 flex items-center justify-center min-h-[300px] max-h-[70vh] p-2 overflow-hidden select-none">
              {previewPhotoLoading && !previewPhotoError && (
                <div className="absolute inset-0 flex flex-col items-center justify-center gap-2 text-slate-400 bg-slate-950/90 z-10">
                  <i className="fa-solid fa-spinner fa-spin text-2xl text-brand-500"></i>
                  <span className="text-xs font-semibold text-slate-300">Memuat foto...</span>
                </div>
              )}

              <img
                src={previewPhotoUrl}
                alt="Foto Presensi"
                onLoad={() => {
                  setPreviewPhotoLoading(false);
                  setPreviewPhotoError(false);
                }}
                onError={() => {
                  setPreviewPhotoLoading(false);
                  setPreviewPhotoError(true);
                }}
                className={`max-h-[66vh] max-w-full w-auto h-auto object-contain rounded-2xl shadow-lg transition-all ${
                  previewPhotoLoading || previewPhotoError ? 'hidden' : 'block'
                }`}
              />

              {previewPhotoError && (
                <div className="flex flex-col items-center justify-center text-center p-6 text-slate-400 gap-2 max-w-sm">
                  <i className="fa-solid fa-triangle-exclamation text-amber-400 text-3xl mb-1"></i>
                  <span className="text-xs font-bold text-slate-200">Pratinjau langsung tidak dapat dimuat</span>
                  <p className="text-[11px] text-slate-400">
                    File foto disimpan di Google Drive. Silakan klik tombol di bawah untuk melihat foto secara langsung.
                  </p>
                  <a
                    href={previewPhotoUrl}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="mt-2 px-4 py-2 bg-brand-600 hover:bg-brand-700 text-white rounded-xl text-xs font-bold transition-all inline-flex items-center gap-1.5 shadow-md"
                  >
                    <i className="fa-solid fa-arrow-up-right-from-square"></i> Buka Foto di Google Drive
                  </a>
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* MODAL 8: EDIT REKAP PRESENSI HARIAN */}
      {modalEditRekapOpen && editingRekapItem && (
        <div className="fixed inset-0 z-50 bg-black/70 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-white max-w-md w-full rounded-3xl overflow-hidden shadow-2xl flex flex-col max-h-[92vh]">
            <div className="px-5 py-3.5 bg-slate-900 text-white flex items-center justify-between">
              <div>
                <h3 className="font-extrabold text-sm">Edit Presensi Karyawan</h3>
                <p className="text-[10px] text-slate-400 font-medium">Ubah jam masuk, jam keluar, shift, dan status</p>
              </div>
              <button onClick={() => setModalEditRekapOpen(false)} className="text-slate-400 hover:text-white p-1 transition-all cursor-pointer">
                <i className="fa-solid fa-xmark text-lg"></i>
              </button>
            </div>

            <form
              onSubmit={async e => {
                e.preventDefault();
                setEditRekapSaving(true);
                try {
                  const tgl = filterRekapTanggal || getLocalISODate();
                  const item = editingRekapItem;

                  if (editRekapStatusMasuk !== 'Belum Masuk' && editRekapJamMasuk.trim()) {
                    const masukDateIso = new Date(`${tgl}T${editRekapJamMasuk.trim()}:00`).toISOString();
                    if (item.masuk?.id) {
                      await supabaseClient.from('presensi_presensi').update({
                        shift_nama: editRekapShift,
                        status: editRekapStatusMasuk,
                        created_at: masukDateIso,
                        keterangan: editRekapKeterangan.trim() || null
                      }).eq('id', item.masuk.id);
                    } else {
                      await supabaseClient.from('presensi_presensi').insert([{
                        karyawan_id: item.karyawan_id,
                        karyawan_nama: item.karyawan_nama,
                        divisi: item.divisi,
                        tipe: 'masuk',
                        status: editRekapStatusMasuk,
                        shift_nama: editRekapShift,
                        lokasi_nama: item.lokasi_nama !== '-' ? item.lokasi_nama : 'Kantor',
                        tanggal: tgl,
                        created_at: masukDateIso,
                        keterangan: editRekapKeterangan.trim() || null
                      }]);
                    }
                  } else if (editRekapStatusMasuk === 'Belum Masuk' && item.masuk?.id) {
                    await supabaseClient.from('presensi_presensi').delete().eq('id', item.masuk.id);
                  }

                  if (editRekapStatusKeluar !== 'Belum Keluar' && editRekapJamKeluar.trim()) {
                    const keluarDateIso = new Date(`${tgl}T${editRekapJamKeluar.trim()}:00`).toISOString();
                    if (item.keluar?.id) {
                      await supabaseClient.from('presensi_presensi').update({
                        status: editRekapStatusKeluar,
                        shift_nama: editRekapShift,
                        created_at: keluarDateIso,
                        keterangan: editRekapKeterangan.trim() || null
                      }).eq('id', item.keluar.id);
                    } else {
                      await supabaseClient.from('presensi_presensi').insert([{
                        karyawan_id: item.karyawan_id,
                        karyawan_nama: item.karyawan_nama,
                        divisi: item.divisi,
                        tipe: 'keluar',
                        status: editRekapStatusKeluar,
                        shift_nama: editRekapShift,
                        lokasi_nama: item.lokasi_nama !== '-' ? item.lokasi_nama : 'Kantor',
                        tanggal: tgl,
                        created_at: keluarDateIso,
                        keterangan: editRekapKeterangan.trim() || null
                      }]);
                    }
                  } else if (editRekapStatusKeluar === 'Belum Keluar' && item.keluar?.id) {
                    await supabaseClient.from('presensi_presensi').delete().eq('id', item.keluar.id);
                  }

                  showToast('Data presensi berhasil diperbarui!', 'success');
                  setModalEditRekapOpen(false);
                  loadAdminRekapan();
                } catch {
                  showToast('Gagal memperbarui presensi.', 'error');
                } finally {
                  setEditRekapSaving(false);
                }
              }}
              className="p-5 overflow-y-auto space-y-3.5 text-xs"
            >
              <div className="p-3 bg-slate-50 border border-slate-200 rounded-2xl flex items-center justify-between">
                <div>
                  <div className="text-[10px] font-bold text-slate-400 uppercase">Karyawan</div>
                  <div className="font-bold text-slate-900 text-sm">{editingRekapItem.karyawan_nama}</div>
                  <div className="text-[11px] text-brand-600 font-semibold mt-0.5">{editingRekapItem.divisi}</div>
                </div>
                <div className="text-right">
                  <div className="text-[10px] font-bold text-slate-400 uppercase">Tanggal</div>
                  <div className="font-mono font-bold text-slate-700">{filterRekapTanggal || getLocalISODate()}</div>
                </div>
              </div>

              <div>
                <label className="block font-bold text-slate-700 uppercase mb-1 text-[11px]">Shift Penugasan</label>
                <select
                  value={editRekapShift}
                  onChange={e => setEditRekapShift(e.target.value)}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl font-semibold text-xs focus:ring-2 focus:ring-brand-500"
                >
                  {adminShiftList.map(s => (
                    <option key={s.id} value={s.nama_shift}>{s.nama_shift}</option>
                  ))}
                  {!adminShiftList.some(s => s.nama_shift === editRekapShift) && editRekapShift && (
                    <option value={editRekapShift}>{editRekapShift}</option>
                  )}
                </select>
              </div>

              {/* Edit Jam Masuk */}
              <div className="p-3 bg-emerald-50/50 border border-emerald-100 rounded-2xl space-y-2">
                <div className="font-bold text-emerald-800 flex items-center gap-1.5 text-xs">
                  <i className="fa-solid fa-arrow-right-to-bracket text-emerald-600"></i> Presensi Masuk
                </div>
                <div className="grid grid-cols-2 gap-2">
                  <div>
                    <label className="block text-[10px] font-bold text-slate-600 uppercase mb-1">Jam Masuk (JJ:MM)</label>
                    <input
                      type="text"
                      value={editRekapJamMasuk}
                      onChange={e => setEditRekapJamMasuk(formatAuto24HourTime(e.target.value))}
                      placeholder="08:00"
                      className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl font-mono font-bold text-center text-xs"
                    />
                  </div>
                  <div>
                    <label className="block text-[10px] font-bold text-slate-600 uppercase mb-1">Status Masuk</label>
                    <select
                      value={editRekapStatusMasuk}
                      onChange={e => setEditRekapStatusMasuk(e.target.value)}
                      className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl font-semibold text-xs"
                    >
                      <option value="Hadir">Hadir</option>
                      <option value="Terlambat">Terlambat</option>
                      <option value="Belum Masuk">Belum Masuk / Hapus</option>
                    </select>
                  </div>
                </div>
              </div>

              {/* Edit Jam Keluar */}
              <div className="p-3 bg-rose-50/50 border border-rose-100 rounded-2xl space-y-2">
                <div className="font-bold text-rose-800 flex items-center gap-1.5 text-xs">
                  <i className="fa-solid fa-arrow-right-from-bracket text-rose-600"></i> Presensi Keluar / Status
                </div>
                <div className="grid grid-cols-2 gap-2">
                  <div>
                    <label className="block text-[10px] font-bold text-slate-600 uppercase mb-1">Jam Keluar (JJ:MM)</label>
                    <input
                      type="text"
                      value={editRekapJamKeluar}
                      onChange={e => setEditRekapJamKeluar(formatAuto24HourTime(e.target.value))}
                      placeholder="17:00"
                      className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl font-mono font-bold text-center text-xs"
                    />
                  </div>
                  <div>
                    <label className="block text-[10px] font-bold text-slate-600 uppercase mb-1">Status Keluar</label>
                    <select
                      value={editRekapStatusKeluar}
                      onChange={e => setEditRekapStatusKeluar(e.target.value)}
                      className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl font-semibold text-xs"
                    >
                      <option value="PULANG">PULANG</option>
                      <option value="IZIN">IZIN</option>
                      <option value="SAKIT">SAKIT</option>
                      <option value="LIBUR">LIBUR</option>
                      <option value="CUTI">CUTI</option>
                      <option value="Belum Keluar">Belum keluar</option>
                    </select>
                  </div>
                </div>
              </div>

              <div>
                <label className="block font-bold text-slate-700 uppercase mb-1 text-[11px]">Keterangan Catatan</label>
                <textarea
                  value={editRekapKeterangan}
                  onChange={e => setEditRekapKeterangan(e.target.value)}
                  rows={2}
                  placeholder="Catatan tambahan (opsional)..."
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl font-medium text-xs"
                ></textarea>
              </div>

              <div className="pt-2 flex items-center justify-end gap-2 border-t border-slate-100">
                <button type="button" onClick={() => setModalEditRekapOpen(false)} className="px-4 py-2 text-xs font-bold text-slate-600 hover:bg-slate-100 rounded-xl cursor-pointer">
                  Batal
                </button>
                <button
                  type="submit"
                  disabled={editRekapSaving}
                  className="px-5 py-2 text-xs font-bold bg-brand-600 hover:bg-brand-700 text-white rounded-xl shadow-md cursor-pointer disabled:opacity-50"
                >
                  {editRekapSaving ? 'Menyimpan...' : 'Simpan Perubahan'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* MODAL 9: KARYAWAN FORM */}
      {modalKaryawanOpen && (
        <div className="fixed inset-0 z-50 bg-black/70 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-white max-w-xl w-full rounded-3xl overflow-hidden shadow-2xl flex flex-col max-h-[90vh]">
            <div className="px-5 py-3.5 bg-slate-900 text-white flex items-center justify-between">
              <h3 className="font-extrabold text-sm">{editingKaryawanId ? 'Edit Master Karyawan' : 'Tambah Master Karyawan'}</h3>
              <button onClick={() => setModalKaryawanOpen(false)} className="text-slate-400 hover:text-white cursor-pointer">
                <i className="fa-solid fa-xmark text-lg"></i>
              </button>
            </div>

            <form
              onSubmit={async e => {
                e.preventDefault();
                const matchedDiv = adminDivisiList.find(d => d.nama_divisi === karyawanFormDivisi);
                const connectedShifts = matchedDiv?.shifts && matchedDiv.shifts.length > 0 ? matchedDiv.shifts : ['Reguler'];

                const selProvObj = provinsiList.find(p => p.id === selectedProvId);
                const selKotaObj = kotaList.find(k => k.id === selectedKotaId);
                const selKecObj = kecamatanList.find(kc => kc.id === selectedKecId);
                const selDesaObj = desaList.find(ds => ds.id === selectedDesaId);

                const addrParts = [
                  karyawanFormJalan.trim(),
                  selDesaObj?.name ? `Desa/Kel. ${selDesaObj.name}` : '',
                  selKecObj?.name ? `Kec. ${selKecObj.name}` : '',
                  selKotaObj?.name,
                  selProvObj?.name
                ].filter(Boolean);
                const readableFull = addrParts.join(', ');

                const addressPayload = {
                  jalan: karyawanFormJalan.trim(),
                  provId: selectedProvId || '',
                  provName: selProvObj?.name || '',
                  kotaId: selectedKotaId || '',
                  kotaName: selKotaObj?.name || '',
                  kecId: selectedKecId || '',
                  kecName: selKecObj?.name || '',
                  desaId: selectedDesaId || '',
                  desaName: selDesaObj?.name || '',
                  fullAddress: readableFull || karyawanFormJalan.trim()
                };
                const isoBirth = ddmmyyyyToIso(karyawanFormTglLahir);

                const payload: any = {
                  nama_lengkap: karyawanFormNama.trim(),
                  tgl_lahir: isoBirth || null,
                  username: karyawanFormUsername.trim(),
                  no_wa: karyawanFormWa.trim(),
                  email: karyawanFormEmail.trim(),
                  alamat: JSON.stringify(addressPayload),
                  divisi: karyawanFormDivisi,
                  shifts: connectedShifts
                };

                if (karyawanFormPass.trim()) {
                  payload.password = karyawanFormPass.trim();
                }

                if (editingKaryawanId) {
                  await supabaseClient.from('karyawan_presensi').update(payload).eq('id', editingKaryawanId);
                } else {
                  payload.password = karyawanFormPass.trim() || '123456';
                  await supabaseClient.from('karyawan_presensi').insert([payload]);
                }

                showToast('Data Karyawan Tersimpan!', 'success');
                setModalKaryawanOpen(false);
                loadAdminKaryawan();
              }}
              className="p-5 overflow-y-auto space-y-4"
            >
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Nama Lengkap</label>
                  <input
                    type="text"
                    value={karyawanFormNama}
                    onChange={e => setKaryawanFormNama(e.target.value)}
                    required
                    placeholder="Nama lengkap karyawan"
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-medium"
                  />
                </div>
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Tanggal Lahir (DD/MM/YYYY)</label>
                  <input
                    type="text"
                    value={karyawanFormTglLahir}
                    onChange={e => setKaryawanFormTglLahir(formatAutoDateInput(e.target.value))}
                    required
                    placeholder="Contoh: 15/08/1998"
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-semibold"
                  />
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Username Login</label>
                  <input
                    type="text"
                    value={karyawanFormUsername}
                    onChange={e => setKaryawanFormUsername(e.target.value)}
                    required
                    placeholder="username"
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-medium"
                  />
                </div>
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">No. WhatsApp</label>
                  <input
                    type="text"
                    value={karyawanFormWa}
                    onChange={e => setKaryawanFormWa(e.target.value)}
                    required
                    placeholder="08xxxxxxxx"
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-medium"
                  />
                </div>
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Email</label>
                  <input
                    type="email"
                    value={karyawanFormEmail}
                    onChange={e => setKaryawanFormEmail(e.target.value)}
                    required
                    placeholder="email@perusahaan.com"
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-medium"
                  />
                </div>
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Password Login</label>
                <input
                  type="text"
                  value={karyawanFormPass}
                  onChange={e => setKaryawanFormPass(e.target.value)}
                  placeholder="Kosongkan jika tidak ingin mengubah (default: 123456)"
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-medium"
                />
              </div>

              {/* Divisi & Linked shifts */}
              <div className="p-3 bg-brand-50/50 rounded-2xl border border-brand-100 space-y-2">
                <label className="block text-xs font-bold text-slate-800">Pilih Divisi Karyawan</label>
                <select
                  value={karyawanFormDivisi}
                  onChange={e => setKaryawanFormDivisi(e.target.value)}
                  className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs font-semibold"
                >
                  {adminDivisiList.map(d => (
                    <option key={d.id} value={d.nama_divisi}>{d.nama_divisi}</option>
                  ))}
                </select>
                <div className="flex items-center gap-1.5 text-[11px] text-slate-500">
                  <span className="font-bold text-brand-700">Shift Terhubung:</span>
                  <div className="flex flex-wrap gap-1">
                    {(adminDivisiList.find(d => d.nama_divisi === karyawanFormDivisi)?.shifts || ['Reguler']).map(s => (
                      <span key={s} className="px-2 py-0.5 rounded-lg bg-indigo-50 text-indigo-700 text-[11px] font-bold border border-indigo-100 flex items-center gap-1">
                        <i className="fa-solid fa-clock text-[10px]"></i> {s}
                      </span>
                    ))}
                  </div>
                </div>
              </div>

              {/* Indonesian Address Cascade */}
              <div className="p-3 bg-slate-50 rounded-2xl border border-slate-200 space-y-3">
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-1.5">
                    <span className="text-xs font-bold text-slate-800 uppercase">Alamat Lengkap Karyawan</span>
                    {regionLoading && (
                      <span className="text-[10px] font-bold text-brand-600 animate-pulse">
                        <i className="fa-solid fa-spinner fa-spin mr-1"></i> Memuat wilayah...
                      </span>
                    )}
                  </div>
                  {(selectedProvId || selectedKotaId || selectedKecId || selectedDesaId) && (
                    <button
                      type="button"
                      onClick={() => {
                        setSelectedProvId('');
                        setSelectedKotaId('');
                        setSelectedKecId('');
                        setSelectedDesaId('');
                        setKotaList([]);
                        setKecamatanList([]);
                        setDesaList([]);
                      }}
                      className="text-[10px] font-bold text-rose-600 hover:text-rose-700 transition-colors flex items-center gap-1 cursor-pointer"
                      title="Hapus pilihan wilayah"
                    >
                      <i className="fa-solid fa-rotate-left text-[9px]"></i> Reset Wilayah
                    </button>
                  )}
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                  <div>
                    <label className="block text-[11px] text-slate-500 font-semibold mb-1">Provinsi</label>
                    <select
                      value={selectedProvId}
                      onChange={e => handleProvinsiChange(e.target.value)}
                      className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs"
                    >
                      <option value="">Pilih Provinsi</option>
                      {provinsiList.map(pr => <option key={pr.id} value={pr.id}>{pr.name}</option>)}
                    </select>
                  </div>
                  <div>
                    <label className="block text-[11px] text-slate-500 font-semibold mb-1">Kabupaten / Kota</label>
                    <select
                      value={selectedKotaId}
                      onChange={e => handleKotaChange(e.target.value)}
                      disabled={!selectedProvId}
                      className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs disabled:bg-slate-100 disabled:text-slate-400"
                    >
                      <option value="">Pilih Kabupaten / Kota</option>
                      {kotaList.map(k => <option key={k.id} value={k.id}>{k.name}</option>)}
                    </select>
                  </div>
                  <div>
                    <label className="block text-[11px] text-slate-500 font-semibold mb-1">Kecamatan</label>
                    <select
                      value={selectedKecId}
                      onChange={e => handleKecamatanChange(e.target.value)}
                      disabled={!selectedKotaId}
                      className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs disabled:bg-slate-100 disabled:text-slate-400"
                    >
                      <option value="">Pilih Kecamatan</option>
                      {kecamatanList.map(kc => <option key={kc.id} value={kc.id}>{kc.name}</option>)}
                    </select>
                  </div>
                  <div>
                    <label className="block text-[11px] text-slate-500 font-semibold mb-1">Desa / Kelurahan</label>
                    <select
                      value={selectedDesaId}
                      onChange={e => setSelectedDesaId(e.target.value)}
                      disabled={!selectedKecId}
                      className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs disabled:bg-slate-100 disabled:text-slate-400"
                    >
                      <option value="">Pilih Desa / Kelurahan</option>
                      {desaList.map(ds => <option key={ds.id} value={ds.id}>{ds.name}</option>)}
                    </select>
                  </div>
                </div>

                <div>
                  <label className="block text-[11px] text-slate-500 font-semibold mb-1">Detail Jalan / RT-RW / No. Rumah</label>
                  <input
                    type="text"
                    value={karyawanFormJalan}
                    onChange={e => setKaryawanFormJalan(e.target.value)}
                    placeholder="Contoh: Jl. Diponegoro No. 12 RT 02 / RW 01"
                    className="w-full px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs font-medium"
                  />
                </div>

                {/* Live Preview Alamat Lengkap */}
                <div className="p-2.5 bg-brand-50/70 border border-brand-100 rounded-xl space-y-1">
                  <div className="flex items-center gap-1.5 text-[10px] font-extrabold text-brand-700 uppercase">
                    <i className="fa-solid fa-map-pin"></i>
                    <span>Pratinjau Alamat Lengkap Tersimpan:</span>
                  </div>
                  <p className="text-[11px] font-bold text-slate-800 leading-relaxed">
                    {[
                      karyawanFormJalan.trim(),
                      desaList.find(ds => ds.id === selectedDesaId)?.name ? `Desa/Kel. ${desaList.find(ds => ds.id === selectedDesaId)?.name}` : '',
                      kecamatanList.find(kc => kc.id === selectedKecId)?.name ? `Kec. ${kecamatanList.find(kc => kc.id === selectedKecId)?.name}` : '',
                      kotaList.find(k => k.id === selectedKotaId)?.name,
                      provinsiList.find(p => p.id === selectedProvId)?.name
                    ].filter(Boolean).join(', ') || '(Belum mengisi rincian alamat)'}
                  </p>
                </div>
              </div>

              <div className="pt-2 flex justify-end gap-2">
                <button type="button" onClick={() => setModalKaryawanOpen(false)} className="px-4 py-2 text-xs font-bold text-slate-600 hover:bg-slate-100 rounded-xl cursor-pointer">
                  Batal
                </button>
                <button type="submit" className="px-5 py-2 text-xs font-bold bg-brand-600 hover:bg-brand-700 text-white rounded-xl shadow-md cursor-pointer">
                  Simpan Data Karyawan
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* MODAL 10: LOKASI FORM */}
      {modalLokasiOpen && (
        <div className="fixed inset-0 z-50 bg-black/70 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-white max-w-sm w-full rounded-3xl p-5 shadow-2xl space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-2">
              <h3 className="font-extrabold text-sm text-slate-900">{editingLokasiId ? 'Edit Lokasi Kantor' : 'Tambah Lokasi Kantor'}</h3>
              <button onClick={() => setModalLokasiOpen(false)} className="text-slate-400 hover:text-slate-700 cursor-pointer">
                <i className="fa-solid fa-xmark"></i>
              </button>
            </div>
            <form
              onSubmit={async e => {
                e.preventDefault();
                const payload = {
                  nama_lokasi: lokasiNama.trim(),
                  latitude: parseFloat(lokasiLat),
                  longitude: parseFloat(lokasiLng),
                  radius_meter: parseInt(lokasiRadius, 10) || 100
                };
                if (editingLokasiId) {
                  await supabaseClient.from('lokasi_presensi').update(payload).eq('id', editingLokasiId);
                } else {
                  await supabaseClient.from('lokasi_presensi').insert([payload]);
                }
                showToast('Lokasi berhasil disimpan!', 'success');
                setModalLokasiOpen(false);
                loadAdminLokasi();
              }}
              className="space-y-3"
            >
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Nama Lokasi</label>
                <input
                  type="text"
                  required
                  value={lokasiNama}
                  onChange={e => setLokasiNama(e.target.value)}
                  placeholder="Contoh: Kantor Cabang Barat"
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-medium"
                />
              </div>
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Latitude</label>
                  <input
                    type="number"
                    step="any"
                    required
                    value={lokasiLat}
                    onChange={e => setLokasiLat(e.target.value)}
                    placeholder="-6.9932"
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-semibold"
                  />
                </div>
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Longitude</label>
                  <input
                    type="number"
                    step="any"
                    required
                    value={lokasiLng}
                    onChange={e => setLokasiLng(e.target.value)}
                    placeholder="110.4203"
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-semibold"
                  />
                </div>
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Radius Presensi (Meter)</label>
                <input
                  type="number"
                  required
                  value={lokasiRadius}
                  onChange={e => setLokasiRadius(e.target.value)}
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-bold"
                />
              </div>
              <button type="submit" className="w-full py-2.5 bg-brand-600 hover:bg-brand-700 text-white rounded-xl text-xs font-bold shadow-md cursor-pointer">
                Simpan Lokasi
              </button>
            </form>
          </div>
        </div>
      )}

      {/* MODAL 11: SHIFT FORM */}
      {modalShiftOpen && (
        <div className="fixed inset-0 z-50 bg-black/70 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-white max-w-sm w-full rounded-3xl p-5 shadow-2xl space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-2">
              <h3 className="font-extrabold text-sm text-slate-900">{editingShiftId ? 'Edit Shift Kerja' : 'Tambah Shift Kerja'}</h3>
              <button onClick={() => setModalShiftOpen(false)} className="text-slate-400 hover:text-slate-700 cursor-pointer">
                <i className="fa-solid fa-xmark"></i>
              </button>
            </div>
            <form
              onSubmit={async e => {
                e.preventDefault();
                const jm = shiftJamMasuk.trim();
                const jp = shiftJamPulang.trim();
                const payload = {
                  nama_shift: shiftNama.trim(),
                  jam_masuk: jm.length === 5 ? jm + ':00' : jm,
                  jam_pulang: jp.length === 5 ? jp + ':00' : jp
                };
                if (editingShiftId) {
                  await supabaseClient.from('shift_presensi').update(payload).eq('id', editingShiftId);
                } else {
                  await supabaseClient.from('shift_presensi').insert([payload]);
                }
                showToast('Shift berhasil disimpan!', 'success');
                setModalShiftOpen(false);
                loadAdminShiftAndJam();
              }}
              className="space-y-3"
            >
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Nama Shift</label>
                <input
                  type="text"
                  required
                  value={shiftNama}
                  onChange={e => setShiftNama(e.target.value)}
                  placeholder="Contoh: Shift Pagi"
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-medium"
                />
              </div>
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Jam Masuk (JJ:MM)</label>
                  <input
                    type="text"
                    required
                    value={shiftJamMasuk}
                    onChange={e => setShiftJamMasuk(formatAuto24HourTime(e.target.value))}
                    placeholder="08:00"
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-mono font-bold text-center"
                  />
                </div>
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">Jam Pulang (JJ:MM)</label>
                  <input
                    type="text"
                    required
                    value={shiftJamPulang}
                    onChange={e => setShiftJamPulang(formatAuto24HourTime(e.target.value))}
                    placeholder="16:00"
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-mono font-bold text-center"
                  />
                </div>
              </div>
              <p className="text-[10px] text-slate-400 italic">Cukup ketik 4 angka, tanda titik dua (:) akan otomatis terpasang.</p>
              <button type="submit" className="w-full py-2.5 bg-brand-600 hover:bg-brand-700 text-white rounded-xl text-xs font-bold shadow-md cursor-pointer">
                Simpan Shift
              </button>
            </form>
          </div>
        </div>
      )}

      {/* MODAL 12: DIVISI FORM */}
      {modalDivisiOpen && (
        <div className="fixed inset-0 z-50 bg-black/70 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-white max-w-sm w-full rounded-3xl p-5 shadow-2xl space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-2">
              <h3 className="font-extrabold text-sm text-slate-900">{editingDivisiId ? 'Edit Divisi' : 'Tambah Divisi'}</h3>
              <button onClick={() => setModalDivisiOpen(false)} className="text-slate-400 hover:text-slate-700 cursor-pointer">
                <i className="fa-solid fa-xmark"></i>
              </button>
            </div>
            <form
              onSubmit={async e => {
                e.preventDefault();
                const payload = {
                  nama_divisi: divisiNama.trim(),
                  shifts: divisiSelectedShifts,
                  allow_cuti: divisiAllowCuti
                };
                if (editingDivisiId) {
                  await supabaseClient.from('divisi_presensi').update(payload).eq('id', editingDivisiId);
                } else {
                  await supabaseClient.from('divisi_presensi').insert([payload]);
                }
                showToast('Divisi berhasil disimpan!', 'success');
                setModalDivisiOpen(false);
                loadAdminDivisi();
              }}
              className="space-y-3"
            >
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Nama Divisi</label>
                <input
                  type="text"
                  required
                  value={divisiNama}
                  onChange={e => setDivisiNama(e.target.value)}
                  placeholder="Contoh: Beauty Advisor (BA)"
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-medium"
                />
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Tautkan Shift Kerja</label>
                <div className="max-h-36 overflow-y-auto space-y-1.5 p-2.5 bg-slate-50 border border-slate-200 rounded-xl">
                  {adminShiftList.map(s => (
                    <label key={s.id} className="flex items-center gap-2 text-xs font-semibold text-slate-700 cursor-pointer">
                      <input
                        type="checkbox"
                        checked={divisiSelectedShifts.includes(s.nama_shift)}
                        onChange={e => {
                          if (e.target.checked) {
                            setDivisiSelectedShifts(prev => [...prev, s.nama_shift]);
                          } else {
                            setDivisiSelectedShifts(prev => prev.filter(x => x !== s.nama_shift));
                          }
                        }}
                        className="rounded text-brand-600"
                      />
                      <span>{s.nama_shift}</span>
                    </label>
                  ))}
                </div>
              </div>
              <div className="flex items-center gap-2 pt-1">
                <input
                  type="checkbox"
                  id="divisiAllowCutiCheckbox"
                  checked={divisiAllowCuti}
                  onChange={e => setDivisiAllowCuti(e.target.checked)}
                  className="w-4 h-4 text-brand-600 rounded cursor-pointer"
                />
                <label htmlFor="divisiAllowCutiCheckbox" className="text-xs font-semibold text-slate-700 cursor-pointer">
                  Aktifkan tombol cuti untuk divisi ini
                </label>
              </div>
              <button type="submit" className="w-full py-2.5 bg-brand-600 hover:bg-brand-700 text-white rounded-xl text-xs font-bold shadow-md cursor-pointer">
                Simpan Divisi
              </button>
            </form>
          </div>
        </div>
      )}

      {/* MODAL 13: ADMIN USER FORM */}
      {modalAdminOpen && (
        <div className="fixed inset-0 z-50 bg-black/70 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-white max-w-sm w-full rounded-3xl p-5 shadow-2xl space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-2">
              <h3 className="font-extrabold text-sm text-slate-900">{editingAdminId ? 'Edit Akun Admin' : 'Tambah Akun Admin'}</h3>
              <button onClick={() => setModalAdminOpen(false)} className="text-slate-400 hover:text-slate-700 cursor-pointer">
                <i className="fa-solid fa-xmark"></i>
              </button>
            </div>
            <form
              onSubmit={async e => {
                e.preventDefault();
                const payload: any = {
                  nama_lengkap: adminUserNama.trim(),
                  username: adminUserUsername.trim(),
                  email: adminUserEmail.trim()
                };
                if (adminUserPass.trim()) {
                  payload.password = adminUserPass.trim();
                }
                if (editingAdminId) {
                  await supabaseClient.from('admin_presensi').update(payload).eq('id', editingAdminId);
                } else {
                  payload.password = adminUserPass.trim() || 'admin123';
                  await supabaseClient.from('admin_presensi').insert([payload]);
                }
                showToast('Data Admin Tersimpan!', 'success');
                setModalAdminOpen(false);
                loadAdminUsers();
              }}
              className="space-y-3"
            >
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Nama Lengkap</label>
                <input
                  type="text"
                  required
                  value={adminUserNama}
                  onChange={e => setAdminUserNama(e.target.value)}
                  placeholder="Administrator HRD"
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-medium"
                />
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Username</label>
                <input
                  type="text"
                  required
                  value={adminUserUsername}
                  onChange={e => setAdminUserUsername(e.target.value)}
                  placeholder="admin"
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-medium"
                />
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Email</label>
                <input
                  type="email"
                  required
                  value={adminUserEmail}
                  onChange={e => setAdminUserEmail(e.target.value)}
                  placeholder="admin@perusahaan.com"
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-medium"
                />
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-700 mb-1">Password</label>
                <input
                  type="text"
                  value={adminUserPass}
                  onChange={e => setAdminUserPass(e.target.value)}
                  placeholder="Default: admin123"
                  className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-medium"
                />
              </div>
              <button type="submit" className="w-full py-2.5 bg-brand-600 hover:bg-brand-700 text-white rounded-xl text-xs font-bold shadow-md cursor-pointer">
                Simpan Admin
              </button>
            </form>
          </div>
        </div>
      )}

      {/* MODAL 14: CONFIRM DELETE */}
      {modalConfirmDeleteOpen && (
        <div className="fixed inset-0 z-50 bg-black/70 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-white max-w-sm w-full rounded-3xl p-6 shadow-2xl border border-slate-100 text-center space-y-4">
            <div className="w-14 h-14 rounded-2xl bg-rose-50 text-rose-600 flex items-center justify-center mx-auto shadow-inner text-2xl">
              <i className="fa-solid fa-triangle-exclamation"></i>
            </div>
            <div>
              <h3 className="font-extrabold text-base text-slate-900">Konfirmasi Hapus</h3>
              <p className="text-xs text-slate-500 mt-1">{deleteConfirmMsg}</p>
            </div>
            <div className="grid grid-cols-2 gap-2 pt-2">
              <button
                type="button"
                onClick={() => {
                  setModalConfirmDeleteOpen(false);
                  setDeleteActionFn(null);
                }}
                className="py-2.5 px-4 rounded-xl text-xs font-bold text-slate-600 bg-slate-100 hover:bg-slate-200 transition-all cursor-pointer"
              >
                Batal
              </button>
              <button
                type="button"
                disabled={deleteExecuting}
                onClick={async () => {
                  if (deleteActionFn) {
                    setDeleteExecuting(true);
                    try {
                      await deleteActionFn();
                      setModalConfirmDeleteOpen(false);
                      setDeleteActionFn(null);
                    } catch (err: any) {
                      showToast('Gagal menghapus: ' + (err.message || 'Mungkin masih terikat relasi data.'), 'error');
                    } finally {
                      setDeleteExecuting(false);
                    }
                  }
                }}
                className="py-2.5 px-4 rounded-xl text-xs font-bold text-white bg-rose-600 hover:bg-rose-700 shadow-md shadow-rose-500/20 transition-all cursor-pointer disabled:opacity-50"
              >
                {deleteExecuting ? 'Menghapus...' : 'Hapus Sekarang'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL 15: EXPORT EXCEL DENGAN PEMILIHAN PERIODE TANGGAL */}
      {modalExportExcelOpen && (
        <div className="fixed inset-0 z-50 bg-black/70 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-white max-w-md w-full rounded-3xl p-6 shadow-2xl border border-slate-100 space-y-4">
            <div className="flex items-center justify-between border-b border-slate-100 pb-3">
              <div className="flex items-center gap-2.5">
                <div className="w-10 h-10 rounded-2xl bg-emerald-50 text-emerald-600 flex items-center justify-center text-lg shadow-inner">
                  <i className="fa-solid fa-file-excel"></i>
                </div>
                <div>
                  <h3 className="font-extrabold text-sm text-slate-900">Export Rekap Presensi (.xlsx)</h3>
                  <p className="text-[10px] text-slate-400 font-medium">Pilih rentang tanggal periode yang ingin diekspor</p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setModalExportExcelOpen(false)}
                className="text-slate-400 hover:text-slate-700 text-sm p-1 cursor-pointer"
              >
                <i className="fa-solid fa-xmark"></i>
              </button>
            </div>

            <div className="space-y-3.5">
              {/* Pilihan Cepat Rentang Tanggal (Presets) */}
              <div>
                <span className="block text-[11px] font-bold text-slate-600 uppercase mb-1.5">Pilihan Periode Cepat</span>
                <div className="grid grid-cols-3 gap-1.5">
                  <button
                    type="button"
                    onClick={() => {
                      const today = getLocalISODate();
                      setExportStartDate(today);
                      setExportEndDate(today);
                    }}
                    className="px-2 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-[11px] font-bold transition-all cursor-pointer"
                  >
                    Hari Ini
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      const end = new Date();
                      const start = new Date();
                      start.setDate(end.getDate() - 6);
                      const pad = (n: number) => String(n).padStart(2, '0');
                      setExportStartDate(`${start.getFullYear()}-${pad(start.getMonth() + 1)}-${pad(start.getDate())}`);
                      setExportEndDate(`${end.getFullYear()}-${pad(end.getMonth() + 1)}-${pad(end.getDate())}`);
                    }}
                    className="px-2 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-[11px] font-bold transition-all cursor-pointer"
                  >
                    7 Hari Terakhir
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      const p = getActivePeriod();
                      setExportStartDate(p.startIso);
                      setExportEndDate(p.endIso);
                    }}
                    className="px-2 py-1.5 bg-emerald-50 hover:bg-emerald-100 text-emerald-800 rounded-xl text-[11px] font-bold border border-emerald-200 transition-all cursor-pointer"
                  >
                    Siklus 26 - 25
                  </button>
                </div>
              </div>

              {/* Form Input Tanggal Mulai & Tanggal Selesai */}
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">
                    Tanggal Mulai <span className="text-rose-500">*</span>
                  </label>
                  <input
                    type="date"
                    value={exportStartDate}
                    onChange={e => setExportStartDate(e.target.value)}
                    required
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-semibold focus:outline-none focus:ring-2 focus:ring-emerald-500"
                  />
                </div>
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">
                    Tanggal Selesai <span className="text-rose-500">*</span>
                  </label>
                  <input
                    type="date"
                    value={exportEndDate}
                    onChange={e => setExportEndDate(e.target.value)}
                    required
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs font-semibold focus:outline-none focus:ring-2 focus:ring-emerald-500"
                  />
                </div>
              </div>

              <div className="p-3 bg-emerald-50/60 rounded-2xl border border-emerald-100 text-xs text-emerald-900 flex items-start gap-2">
                <i className="fa-solid fa-circle-info text-emerald-600 mt-0.5 shrink-0"></i>
                <span className="text-[11px] leading-relaxed">
                  File Excel (.xlsx) akan menyertakan ringkasan presensi harian per karyawan beserta jam masuk, jam pulang, status, lokasi, catatan, serta tautan foto bukti presensi.
                </span>
              </div>
            </div>

            <div className="pt-2 flex items-center justify-end gap-2 border-t border-slate-100">
              <button
                type="button"
                onClick={() => setModalExportExcelOpen(false)}
                className="px-4 py-2.5 rounded-xl text-xs font-bold text-slate-600 hover:bg-slate-100 transition-all cursor-pointer"
              >
                Batal
              </button>
              <button
                type="button"
                disabled={exportExcelLoading}
                onClick={handleDownloadExcelPeriod}
                className="px-5 py-2.5 rounded-xl text-xs font-bold bg-emerald-600 hover:bg-emerald-700 text-white shadow-md shadow-emerald-500/25 transition-all flex items-center gap-1.5 cursor-pointer disabled:opacity-50"
              >
                {exportExcelLoading ? (
                  <>
                    <i className="fa-solid fa-spinner fa-spin"></i>
                    <span>Memproses Excel...</span>
                  </>
                ) : (
                  <>
                    <i className="fa-solid fa-download"></i>
                    <span>Download File Excel</span>
                  </>
                )}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL PENGUMUMAN DETAIL */}
      {modalPengumumanDetail && (
        <div className="fixed inset-0 z-50 bg-black/70 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-white max-w-md w-full rounded-3xl overflow-hidden shadow-2xl flex flex-col max-h-[90vh]">
            <div className="px-5 py-4 bg-gradient-to-r from-amber-500 to-orange-500 text-white flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <div className="w-8 h-8 rounded-xl bg-white/20 flex items-center justify-center text-sm font-bold">
                  <i className="fa-solid fa-bullhorn"></i>
                </div>
                <div>
                  <h3 className="font-extrabold text-sm">Pengumuman Kantor</h3>
                  <p className="text-[10px] text-amber-100">Disiarkan oleh manajemen perusahaan</p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setModalPengumumanDetail(null)}
                className="w-8 h-8 rounded-xl bg-white/20 hover:bg-white/30 text-white flex items-center justify-center cursor-pointer"
              >
                <i className="fa-solid fa-xmark text-sm"></i>
              </button>
            </div>

            <div className="p-5 overflow-y-auto space-y-3">
              <div>
                <h4 className="font-black text-base text-slate-900 leading-snug">
                  {modalPengumumanDetail.judul}
                </h4>
                <div className="flex flex-wrap items-center gap-2 mt-2 text-[11px] text-slate-500">
                  <span className="inline-flex items-center gap-1 font-semibold bg-slate-100 text-slate-700 px-2 py-0.5 rounded-md">
                    <i className="fa-solid fa-user-shield text-amber-600"></i>
                    {modalPengumumanDetail.author || 'Admin'}
                  </span>
                  <span className="inline-flex items-center gap-1 font-semibold bg-slate-100 text-slate-700 px-2 py-0.5 rounded-md">
                    <i className="fa-regular fa-calendar text-brand-600"></i>
                    {formatIndoDate(modalPengumumanDetail.created_at?.split('T')[0] || '')} • {formatTime24(modalPengumumanDetail.created_at)} WIB
                  </span>
                </div>
              </div>

              <div className="p-4 bg-slate-50 rounded-2xl border border-slate-100 text-slate-700 text-xs leading-relaxed whitespace-pre-wrap font-medium">
                {modalPengumumanDetail.isi}
              </div>
            </div>

            <div className="p-4 bg-slate-50 border-t border-slate-100 flex items-center justify-end">
              <button
                type="button"
                onClick={() => setModalPengumumanDetail(null)}
                className="px-5 py-2.5 bg-slate-900 hover:bg-black text-white text-xs font-bold rounded-xl cursor-pointer"
              >
                Tutup
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL SEMUA PENGUMUMAN */}
      {modalAllPengumumanOpen && (
        <div className="fixed inset-0 z-50 bg-black/70 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-white max-w-lg w-full rounded-3xl overflow-hidden shadow-2xl flex flex-col max-h-[90vh]">
            <div className="px-5 py-4 bg-slate-900 text-white flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <div className="w-8 h-8 rounded-xl bg-amber-500/20 text-amber-400 flex items-center justify-center text-sm font-bold">
                  <i className="fa-solid fa-bullhorn"></i>
                </div>
                <div>
                  <h3 className="font-extrabold text-sm">Papan Pengumuman Kantor</h3>
                  <p className="text-[10px] text-slate-400">Seluruh siaran pengumuman resmi</p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setModalAllPengumumanOpen(false)}
                className="w-8 h-8 rounded-xl bg-white/10 hover:bg-white/20 text-white flex items-center justify-center cursor-pointer"
              >
                <i className="fa-solid fa-xmark text-sm"></i>
              </button>
            </div>

            <div className="p-5 overflow-y-auto space-y-3">
              {karyawanPengumumanList.length === 0 ? (
                <div className="text-center py-8 text-slate-400 text-xs">
                  Belum ada siaran pengumuman.
                </div>
              ) : (
                karyawanPengumumanList.map(item => (
                  <div
                    key={item.id}
                    onClick={() => {
                      setModalAllPengumumanOpen(false);
                      setModalPengumumanDetail(item);
                    }}
                    className="p-4 bg-slate-50 hover:bg-amber-50/40 border border-slate-200/80 hover:border-amber-200 rounded-2xl transition-all cursor-pointer space-y-2 group shadow-2xs"
                  >
                    <div className="flex items-start justify-between gap-2">
                      <h4 className="font-extrabold text-xs text-slate-900 group-hover:text-brand-600 transition-colors">
                        {item.judul}
                      </h4>
                      <span className="text-[10px] font-semibold text-slate-400 shrink-0">
                        {formatIndoDate(item.created_at?.split('T')[0] || '')}
                      </span>
                    </div>
                    <p className="text-xs text-slate-600 line-clamp-3 leading-relaxed">
                      {item.isi}
                    </p>
                    <div className="flex items-center justify-between text-[10px] text-slate-400 pt-1 border-t border-slate-200/50">
                      <span>Oleh: {item.author || 'Admin'}</span>
                      <span className="text-brand-600 font-bold group-hover:translate-x-0.5 transition-transform flex items-center gap-1">
                        <span>Baca Selengkapnya</span>
                        <i className="fa-solid fa-arrow-right text-[9px]"></i>
                      </span>
                    </div>
                  </div>
                ))
              )}
            </div>

            <div className="p-4 bg-slate-50 border-t border-slate-100 flex items-center justify-end">
              <button
                type="button"
                onClick={() => setModalAllPengumumanOpen(false)}
                className="px-5 py-2.5 bg-slate-900 hover:bg-black text-white text-xs font-bold rounded-xl cursor-pointer"
              >
                Tutup
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Toast Notification Container */}
      <div className="fixed bottom-4 right-4 z-50 flex flex-col gap-2 pointer-events-none">
        {toasts.map(t => {
          const bg =
            t.type === 'success'
              ? 'bg-emerald-600'
              : t.type === 'error'
              ? 'bg-rose-600'
              : t.type === 'warning'
              ? 'bg-amber-600'
              : 'bg-slate-900';
          const icon =
            t.type === 'success'
              ? 'fa-check'
              : t.type === 'error'
              ? 'fa-circle-exclamation'
              : t.type === 'warning'
              ? 'fa-triangle-exclamation'
              : 'fa-info';
          return (
            <div
              key={t.id}
              className={`${bg} text-white text-xs font-semibold px-4 py-3 rounded-2xl shadow-xl flex items-center gap-2 pointer-events-auto transform transition-all duration-300`}
            >
              <i className={`fa-solid ${icon}`}></i>
              <span>{t.message}</span>
            </div>
          );
        })}
      </div>
    </div>
  );
}
