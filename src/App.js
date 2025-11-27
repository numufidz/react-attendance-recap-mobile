import React, { useState } from 'react';
import { Upload, Download, FileText, Loader } from 'lucide-react';
import * as XLSX from 'xlsx';
import { jsPDF } from 'jspdf';
import autoTable from 'jspdf-autotable';
import html2canvas from 'html2canvas'; // Install dulu: npm install html2canvas

const normalizeId = (raw) => {
  const s = String(raw || '').trim();
  const x = s.includes('/') ? s.split('/')[0] : s;
  return x;
};

const AttendanceRecapSystem = () => {
  // Tambahkan CSS untuk mobile
  React.useEffect(() => {
    const style = document.createElement('style');
    style.textContent = `
      @media (max-width: 640px) {
        * {
          -webkit-tap-highlight-color: transparent;
        }
        button, a {
          min-height: 44px;
          min-width: 44px;
        }
        input[type="file"] + label {
          min-height: 100px;
        }
        table {
          font-size: 11px;
        }
        th, td {
          white-space: nowrap;
        }
      }
      
      @media (max-width: 480px) {
        table {
          font-size: 10px;
        }
      }
    `;
    document.head.appendChild(style);
    return () => document.head.removeChild(style);
  }, []);

  const [attendanceData, setAttendanceData] = useState([]);
  const [scheduleData, setScheduleData] = useState([]);
  const [recapData, setRecapData] = useState(null);
  const [summaryData, setSummaryData] = useState(null);
  const [rankingData, setRankingData] = useState(null);
  const [startDate, setStartDate] = useState('2025-11-20');
  const [endDate, setEndDate] = useState('2025-12-05');
  const [isLoadingAttendance, setIsLoadingAttendance] = useState(false);
  const [isLoadingSchedule, setIsLoadingSchedule] = useState(false);
  const [errorMessage, setErrorMessage] = useState('');
  const panduanRef = React.useRef(null);

  // Ref untuk capture kesimpulan sebagai gambar
  const summaryRef = React.useRef(null);

  // Download sebagai JPG
  const handleDownloadSummaryJPG = async () => {
    if (summaryRef.current) {
      const canvas = await html2canvas(summaryRef.current, {
        scale: 2,
        backgroundColor: null,
      });
      const image = canvas.toDataURL('image/jpeg', 1.0);
      const link = document.createElement('a');
      link.href = image;
      link.download = `kesimpulan-absensi-${summaryData.periode}.jpg`;
      link.click();
    }
  };

  // Copy gambar ke clipboard
  const handleCopySummary = async () => {
    if (summaryRef.current) {
      const canvas = await html2canvas(summaryRef.current, {
        scale: 2,
        backgroundColor: null,
      });
      canvas.toBlob(async (blob) => {
        if (blob) {
          await navigator.clipboard.write([
            new ClipboardItem({ 'image/png': blob }),
          ]);
          alert('Gambar kesimpulan berhasil disalin ke clipboard!');
        }
      });
    }
  };

  const handleDownloadPanduan = async () => {
    if (panduanRef.current) {
      const canvas = await html2canvas(panduanRef.current, {
        scale: 2,
        backgroundColor: '#ffffff',
      });
      const image = canvas.toDataURL('image/jpeg', 1.0);
      const link = document.createElement('a');
      link.href = image;
      link.download = 'panduan-rekap.jpg';
      link.click();
    }
  };

  const handleCopyPanduan = async () => {
    if (panduanRef.current) {
      const canvas = await html2canvas(panduanRef.current, {
        scale: 2,
        backgroundColor: '#ffffff',
      });
      canvas.toBlob(async (blob) => {
        if (blob) {
          await navigator.clipboard.write([
            new ClipboardItem({ 'image/png': blob }),
          ]);
          alert('Gambar berhasil disalin ke clipboard!');
        }
      });
    }
  };

  const validateFile = (file) => {
    if (!file) return false;
    const ext = file.name.split('.').pop().toLowerCase();
    if (ext !== 'xlsx' && ext !== 'xls') {
      setErrorMessage('File harus berformat .xlsx atau .xls');
      return false;
    }
    setErrorMessage('');
    return true;
  };

  const processAttendanceFile = (file) => {
    if (!validateFile(file)) return;
    setIsLoadingAttendance(true);

    // Extract tanggal dari nama file
    // Format: attendance_report_detail_2025-11-01_2025-11-22.xlsx
    const fileName = file.name;
    const datePattern = /(\d{4}-\d{2}-\d{2})_(\d{4}-\d{2}-\d{2})/;
    const match = fileName.match(datePattern);

    if (match) {
      const extractedStartDate = match[1]; // 2025-11-01
      const extractedEndDate = match[2]; // 2025-11-22

      // Set tanggal otomatis
      setStartDate(extractedStartDate);
      setEndDate(extractedEndDate);

      console.log(
        '✅ Periode terdeteksi dari nama file:',
        extractedStartDate,
        'sampai',
        extractedEndDate
      );

      // Tampilkan notifikasi
      setTimeout(() => {
        alert(
          `📅 Periode terdeteksi otomatis dari nama file:\n\nDari: ${extractedStartDate}\nSampai: ${extractedEndDate}\n\n✅ Tanggal sudah diset otomatis!`
        );
      }, 1000);
    }

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        const formatTime = (val) => {
          if (!val || val === '-' || val === '') return '-';
          if (typeof val === 'string') {
            if (val.includes(':')) {
              const parts = val.split(':');
              if (parts.length >= 2) {
                const hours = parts[0].padStart(2, '0');
                const minutes = parts[1].padStart(2, '0');
                return hours + ':' + minutes;
              }
            }
            if (val === '-') return '-';
          }
          if (typeof val === 'number') {
            if (val < 0 || val > 1) return '-';
            const totalMinutes = Math.round(val * 24 * 60);
            const hours = Math.floor(totalMinutes / 60);
            const minutes = totalMinutes % 60;
            return (
              String(hours).padStart(2, '0') +
              ':' +
              String(minutes).padStart(2, '0')
            );
          }
          return String(val);
        };

        const attendance = [];
        let currentEmployee = null;

        for (let i = 0; i < jsonData.length; i++) {
          const row = jsonData[i];

          if (row[0] === 'Nama Karyawan' && row[1]) {
            if (currentEmployee && currentEmployee.records.length > 0) {
              attendance.push(currentEmployee);
            }

            currentEmployee = {
              name: row[1] || '',
              id: '',
              position: 'Guru', // default sementara
              records: [],
            };

            // Ambil jabatan dari kolom J pada 3 baris berikutnya
            for (let j = 0; j < 4; j++) {
              if (
                i + j < jsonData.length &&
                jsonData[i + j][9] &&
                typeof jsonData[i + j][9] === 'string'
              ) {
                const val = jsonData[i + j][9].trim();
                if (val && val !== 'MTs. An-Nur Bululawang' && val !== '') {
                  currentEmployee.position = val;
                  break;
                }
              }
            }

            if (i + 1 < jsonData.length && jsonData[i + 1][0] === 'ID/NIK') {
              const idCell = jsonData[i + 1][1] || '';
              currentEmployee.id = normalizeId(idCell);
            }

            if (i + 2 < jsonData.length && jsonData[i + 2][0] === 'Jabatan') {
              // Versi paling aman – cari nilai di kolom setelah "Jabatan"
              let jabatan = '';
              if (row[8] === 'Jabatan' && row[9]) {
                jabatan = row[9];
              } else if (
                (row[9] &&
                  typeof row[9] === 'string' &&
                  row[9].includes('amad')) ||
                row[9].includes('uru')
              ) {
                // fallback langsung ambil kolom J
                jabatan = row[9];
              }
              currentEmployee.position = jabatan.trim();
            }
          }

          if (currentEmployee && row[0]) {
            const cellValue = String(row[0]);
            const dateMatch = cellValue.match(/(\d{4})-(\d{2})-(\d{2})/);
            if (dateMatch) {
              currentEmployee.records.push({
                date: dateMatch[3],
                month: dateMatch[2],
                year: dateMatch[1],
                checkIn: formatTime(row[4]),
                checkOut: formatTime(row[5]),
              });
            }
          }

          if (row[0] === 'TOTAL' && currentEmployee) {
            if (currentEmployee.records.length > 0) {
              attendance.push(currentEmployee);
            }
            currentEmployee = null;
          }
        }

        if (currentEmployee && currentEmployee.records.length > 0) {
          attendance.push(currentEmployee);
        }

        if (attendance.length === 0) {
          throw new Error('Tidak ada data karyawan yang valid ditemukan');
        }

        console.log(
          '✅ Data absensi berhasil dimuat:',
          attendance.length,
          'karyawan'
        );
        setAttendanceData(attendance);
      } catch (error) {
        console.error('Error membaca file absensi:', error);
        setErrorMessage('Gagal membaca file absensi: ' + error.message);
      } finally {
        setIsLoadingAttendance(false);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const processScheduleFile = (file) => {
    if (!validateFile(file)) return;
    setIsLoadingSchedule(true);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        console.log(
          '📊 Raw data jadwal (5 baris pertama):',
          jsonData.slice(0, 5)
        );

        const formatTime = (val) => {
          if (!val || val === 'OFF' || val === 'L' || val === '-') return val;
          if (typeof val === 'string' && val.includes(':')) return val;
          if (typeof val === 'number' && val < 1) {
            const totalMinutes = Math.round(val * 24 * 60);
            const hours = Math.floor(totalMinutes / 60);
            const minutes = totalMinutes % 60;
            return (
              String(hours).padStart(2, '0') +
              ':' +
              String(minutes).padStart(2, '0')
            );
          }
          return String(val);
        };

        // Header ada di baris 1-2, data mulai dari baris 3 (index 2)
        const schedules = jsonData
          .slice(2) // Skip 2 baris header
          .filter((row) => row[0] && row[0] !== '') // ID harus ada
          .map((row) => {
            return {
              id: String(row[0] || '').trim(),
              name: row[2] || '', // Kolom C = NAMA DEPAN
              schedule: {
                sabtu: { start: formatTime(row[3]) }, // Kolom D = SABTU Mulai
                minggu: { start: formatTime(row[5]) }, // Kolom F = AHAD Mulai
                senin: { start: formatTime(row[7]) }, // Kolom H = SENIN Mulai
                selasa: { start: formatTime(row[9]) }, // Kolom J = SELASA Mulai
                rabu: { start: formatTime(row[11]) }, // Kolom L = RABU Mulai
                kamis: { start: formatTime(row[13]) }, // Kolom N = KAMIS Mulai
                jumat: 'L', // JUMAT = Libur
              },
            };
          });

        if (schedules.length === 0) {
          throw new Error('Tidak ada data jadwal yang valid ditemukan');
        }

        console.log(
          '✅ Data jadwal berhasil dimuat:',
          schedules.length,
          'karyawan'
        );
        console.log('📋 Sample jadwal pertama:', schedules[0]);
        setScheduleData(schedules);
      } catch (error) {
        console.error('Error membaca file jadwal:', error);
        setErrorMessage('Gagal membaca file jadwal: ' + error.message);
      } finally {
        setIsLoadingSchedule(false);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const getDateLabel = (date) => {
    const months = [
      'Jan',
      'Feb',
      'Mar',
      'Apr',
      'Mei',
      'Jun',
      'Jul',
      'Ags',
      'Sep',
      'Okt',
      'Nov',
      'Des',
    ];
    const d = new Date(date);
    return d.getDate() + ' ' + months[d.getMonth()];
  };

  const getDayName = (dateStr) => {
    const date = new Date(dateStr);
    const days = [
      'minggu',
      'senin',
      'selasa',
      'rabu',
      'kamis',
      'jumat',
      'sabtu',
    ];
    return days[date.getDay()];
  };

  const timeToMinutes = (timeStr) => {
    if (!timeStr || timeStr === '-' || timeStr === 'OFF' || timeStr === 'L')
      return null;
    const timeString = String(timeStr).trim();
    if (timeString.includes(':')) {
      const parts = timeString.split(':');
      return parseInt(parts[0]) * 60 + parseInt(parts[1]);
    }
    return null;
  };

  const evaluateAttendance = (checkIn, schedule) => {
    if (schedule === 'L' || schedule === 'OFF') {
      return { status: 'L', color: 'FFFFFF', text: 'L' };
    }
    if (schedule === '-') {
      if (checkIn !== '-') {
        return { status: 'H', color: '90EE90', text: 'H' };
      } else {
        return { status: 'A', color: 'FFB3B3', text: '-' };
      }
    }
    if (checkIn === '-') {
      return { status: 'A', color: 'FFB3B3', text: '-' };
    }
    const checkInMin = timeToMinutes(checkIn);
    const schedMin = timeToMinutes(schedule);
    if (checkInMin === null || schedMin === null) {
      return { status: 'H', color: 'FFFF99', text: 'H' };
    }
    if (checkInMin <= schedMin) {
      return { status: 'H', color: '90EE90', text: 'H' };
    } else {
      return { status: 'T', color: 'FFFF99', text: 'H' };
    }
  };

  const getDateRange = () => {
    const start = new Date(startDate);
    const end = new Date(endDate);
    const dates = [];
    for (let d = new Date(start); d <= end; d.setDate(d.getDate() + 1)) {
      dates.push(new Date(d).toISOString().split('T')[0]);
    }
    return dates;
  };

  const generateRecapTables = () => {
    if (attendanceData.length === 0) {
      alert('Silakan upload file laporan absensi');
      return;
    }

    try {
      const dateRange = getDateRange();
      const recap = attendanceData.map((emp, index) => {
        const schedRecord = scheduleData.find((s) => s.id === emp.id);
        const dailyRecords = {};
        const dailyEvaluation = {};

        dateRange.forEach((dateStr) => {
          const [year, month, day] = dateStr.split('-');
          const dayStr = day.padStart(2, '0');
          const record = emp.records.find((r) => {
            const recordDay = r.date.padStart(2, '0');
            return recordDay === dayStr && r.month === month && r.year === year;
          });

          dailyRecords[dateStr] = {
            in: record?.checkIn || '-',
            out: record?.checkOut || '-',
          };

          if (schedRecord) {
            const dayName = getDayName(dateStr);
            const scheduleStart = schedRecord.schedule[dayName]?.start || 'L';
            dailyEvaluation[dateStr] = evaluateAttendance(
              dailyRecords[dateStr].in,
              scheduleStart
            );
          } else {
            const hasIn = dailyRecords[dateStr].in !== '-';
            dailyEvaluation[dateStr] = {
              status: hasIn ? 'H' : 'A',
              color: hasIn ? 'ADD8E6' : 'FFB3B3',
              text: hasIn ? 'H' : '-',
            };
          }
        });

        return {
          no: index + 1,
          id: emp.id,
          name: emp.name,
          position: emp.position,
          dailyRecords,
          dailyEvaluation,
        };
      });
      recap.sort((a, b) => String(a.id).localeCompare(String(b.id)));
      recap.forEach((emp, idx) => {
        emp.no = idx + 1;
      });
      setRecapData({ recap, dateRange });
    } catch (error) {
      alert('Error: ' + error.message);
    }
  };

  const calculateRankings = () => {
    if (!recapData) return null;

    const employeeStats = recapData.recap.map((emp) => {
      const sched = scheduleData.find((s) => s.id === emp.id);
      let hariKerja = 0;
      let hijau = 0; // Disiplin waktu
      let biru = 0; // Tertib administrasi
      let merah = 0; // Alfa

      recapData.dateRange.forEach((dateStr) => {
        const ev = emp.dailyEvaluation[dateStr];
        if (ev.text !== 'L') {
          hariKerja++;
          const rec = emp.dailyRecords[dateStr];
          const hasIn = rec.in !== '-';
          const hasOut = rec.out !== '-';

          // Hitung untuk Tabel 2: Kedisiplinan Waktu
          if (!hasIn && !hasOut) {
            merah++;
          } else if (hasIn && hasOut) {
            biru++;
          } else if (hasIn) {
            const dayName = getDayName(dateStr);
            const schedStart = sched?.schedule[dayName]?.start;
            const inMin = timeToMinutes(rec.in);
            const schedMin = timeToMinutes(schedStart);
            if (schedMin && inMin && inMin <= schedMin) {
              hijau++;
            }
          }
        }
      });

      return {
        id: emp.id,
        name: emp.name,
        position: emp.position,
        hariKerja,
        hijau,
        biru,
        merah,
        persenHijau: hariKerja > 0 ? Math.round((hijau / hariKerja) * 100) : 0,
        persenBiru: hariKerja > 0 ? Math.round((biru / hariKerja) * 100) : 0,
        persenMerah: hariKerja > 0 ? Math.round((merah / hariKerja) * 100) : 0,
      };
    });

    // Top 10 Disiplin Waktu (Hijau persen tertinggi)
    const topDisiplin = [...employeeStats]
      .sort((a, b) => b.persenHijau - a.persenHijau || b.hijau - a.hijau)
      .slice(0, 10);

    // Top 10 Tertib Administrasi (Biru persen tertinggi)
    const topTertib = [...employeeStats]
      .sort((a, b) => b.persenBiru - a.persenBiru || b.biru - a.biru)
      .slice(0, 10);

    // Top 10 Rendah Kesadaran (Merah persen tertinggi)
    const topRendah = [...employeeStats]
      .sort((a, b) => b.persenMerah - a.persenMerah || b.merah - a.merah)
      .slice(0, 10);

    return { topDisiplin, topTertib, topRendah };
  };

  const generateSummary = () => {
    if (!recapData) {
      alert('Silakan generate tabel rekap terlebih dahulu');
      return;
    }

    let totalHariKerja = 0;
    let totalHadir = 0;
    let totalTepat = 0;
    let totalTelat = 0;
    let totalAlfa = 0;

    recapData.recap.forEach((emp) => {
      recapData.dateRange.forEach((dateStr) => {
        const ev = emp.dailyEvaluation[dateStr];
        if (ev.text !== 'L') {
          totalHariKerja++;
          if (ev.text === 'H') {
            totalHadir++;
            if (ev.color === '90EE90') totalTepat++;
            else if (ev.color === 'FFFF99') totalTelat++;
          } else if (ev.text === '-') {
            totalAlfa++;
          }
        }
      });
    });

    const persentaseKehadiran =
      totalHariKerja > 0 ? Math.round((totalHadir / totalHariKerja) * 100) : 0;
    const persentaseTepat =
      totalHariKerja > 0 ? Math.round((totalTepat / totalHariKerja) * 100) : 0;
    const persentaseTelat =
      totalHariKerja > 0 ? Math.round((totalTelat / totalHariKerja) * 100) : 0;
    const persentaseAlfa =
      totalHariKerja > 0 ? Math.round((totalAlfa / totalHariKerja) * 100) : 0;

    let predikat = '';
    let warna = '';
    let icon = '';

    if (persentaseKehadiran >= 96) {
      predikat = 'UNGGUL';
      warna = 'from-green-500 to-emerald-600';
      icon = '🏆';
    } else if (persentaseKehadiran >= 91) {
      predikat = 'BAIK SEKALI / ISTIMEWA';
      warna = 'from-blue-500 to-indigo-600';
      icon = '⭐';
    } else if (persentaseKehadiran >= 86) {
      predikat = 'BAIK';
      warna = 'from-cyan-500 to-blue-600';
      icon = '👍';
    } else if (persentaseKehadiran >= 81) {
      predikat = 'CUKUP';
      warna = 'from-yellow-500 to-orange-600';
      icon = '⚠️';
    } else if (persentaseKehadiran >= 76) {
      predikat = 'BURUK';
      warna = 'from-orange-500 to-red-600';
      icon = '⚡';
    } else {
      predikat = 'BURUK SEKALI';
      warna = 'from-red-500 to-red-700';
      icon = '🚨';
    }

    const startD = new Date(startDate);
    const endD = new Date(endDate);
    const periode = `${startD.toLocaleDateString(
      'id-ID'
    )} - ${endD.toLocaleDateString('id-ID')}`;

    setSummaryData({
      predikat,
      warna,
      icon,
      persentaseKehadiran,
      totalKaryawan: recapData.recap.length,
      totalHariKerja,
      totalHadir,
      persentaseTepat,
      totalTepat,
      persentaseTelat,
      totalTelat,
      persentaseAlfa,
      totalAlfa,
      periode,
    });

    // Hitung ranking
    const rankings = calculateRankings();
    setRankingData(rankings);
  };

  const copyTableToClipboard = (tableId) => {
    const table = document.getElementById(tableId);
    if (!table) return;
    try {
      const range = document.createRange();
      range.selectNode(table);
      const selection = window.getSelection();
      selection.removeAllRanges();
      selection.addRange(range);
      document.execCommand('copy');
      selection.removeAllRanges();
      alert('Tabel berhasil di-copy! Paste di Excel dengan Ctrl+V');
    } catch (err) {
      alert('Gagal copy. Silakan select tabel dan tekan Ctrl+C');
    }
  };

  const downloadTableAsExcel = (tableId, fileName) => {
    const table = document.getElementById(tableId);
    if (!table) return;

    // Clone table untuk manipulasi
    const clonedTable = table.cloneNode(true);

    // Inject inline styles ke setiap cell untuk Excel
    const allCells = clonedTable.querySelectorAll('td, th');
    allCells.forEach((cell) => {
      const classList = cell.className;
      let bgColor = '#FFFFFF';
      let fontWeight = 'normal';
      let textAlign = 'center';

      // Deteksi warna dari className
      if (
        classList.includes('bg-blue-200') ||
        classList.includes('bg-blue-100')
      ) {
        bgColor = '#ADD8E6';
      } else if (classList.includes('bg-blue-300')) {
        bgColor = '#93C5FD';
      } else if (
        classList.includes('bg-yellow-200') ||
        classList.includes('bg-yellow-100')
      ) {
        bgColor = '#FFFF99';
      } else if (classList.includes('bg-yellow-300')) {
        bgColor = '#FDE047';
      } else if (
        classList.includes('bg-red-200') ||
        classList.includes('bg-red-100')
      ) {
        bgColor = '#FFB3B3';
      } else if (
        classList.includes('bg-green-200') ||
        classList.includes('bg-green-100')
      ) {
        bgColor = '#90EE90';
      } else if (classList.includes('bg-green-300')) {
        bgColor = '#86EFAC';
      } else if (classList.includes('bg-gray-100')) {
        bgColor = '#F3F4F6';
      } else if (classList.includes('bg-gray-600')) {
        bgColor = '#4B5563';
      } else if (classList.includes('bg-gray-700')) {
        bgColor = '#374151';
      } else if (classList.includes('bg-purple-100')) {
        bgColor = '#E9D5FF';
      } else if (classList.includes('bg-indigo-100')) {
        bgColor = '#C7D2FE';
      } else if (classList.includes('bg-gray-300')) {
        bgColor = '#D1D5DB';
      }

      // Check inline background color
      if (cell.style.backgroundColor) {
        const inlineColor = cell.style.backgroundColor;
        if (inlineColor.startsWith('rgb')) {
          const match = inlineColor.match(/\d+/g);
          if (match && match.length >= 3) {
            bgColor =
              '#' +
              match
                .slice(0, 3)
                .map((x) => parseInt(x).toString(16).padStart(2, '0'))
                .join('');
          }
        } else if (inlineColor.startsWith('#')) {
          bgColor = inlineColor;
        }
      }

      // Deteksi bold
      if (classList.includes('font-bold')) {
        fontWeight = 'bold';
      }

      // Set inline style untuk Excel
      cell.setAttribute(
        'style',
        `background-color: ${bgColor}; font-weight: ${fontWeight}; text-align: ${textAlign}; border: 1px solid #000000; padding: 5px;`
      );
    });

    // Convert ke HTML string
    let html = '<html><head><meta charset="utf-8"></head><body>';
    html +=
      '<table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse;">';
    html += clonedTable.innerHTML;
    html += '</table></body></html>';

    // Create blob dan download sebagai .xls (HTML format)
    const blob = new Blob([html], {
      type: 'application/vnd.ms-excel;charset=utf-8',
    });

    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `${fileName}.xls`;
    link.click();
    URL.revokeObjectURL(link.href);
  };

  const downloadCompletePdf = async () => {
    if (!summaryData || !recapData) {
      alert('Silakan generate Kesimpulan Profil terlebih dahulu');
      return;
    }

    const doc = new jsPDF('l', 'pt', 'a4'); // LANDSCAPE
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();

    // ============= HALAMAN 1: KESIMPULAN =============
    let yPos = 50;

    // Header
    const headerHeight = 80;
    doc.setFillColor(79, 70, 229);
    doc.rect(0, 0, pageWidth, headerHeight, 'F');
    doc.setTextColor(255, 255, 255);

    // Hitung posisi tengah vertikal
    const centerY = headerHeight / 2;

    // Judul utama
    doc.setFontSize(24);
    doc.setFont(undefined, 'bold');
    doc.text('KESIMPULAN PROFIL ABSENSI', 40, centerY - 8);

    // Subjudul
    doc.setFontSize(16);
    doc.setFont(undefined, 'bold');
    doc.text('MTs. AN-NUR BULULAWANG', 40, centerY + 10);

    // Periode
    doc.setFontSize(12);
    doc.setFont(undefined, 'normal');
    doc.text(`Periode: ${summaryData.periode}`, 40, centerY + 26);

    // yPos setelah header
    yPos = headerHeight + 20;
    doc.setTextColor(0, 0, 0);

    // 1 BARIS 3 KONTAINER: Predikat, Total Karyawan, Total Hari Kerja
    const containerWidth = (pageWidth - 120) / 3;
    const containerHeight = 80;

    // Container 1: Predikat
    doc.setFillColor(240, 240, 255);
    doc.roundedRect(40, yPos, containerWidth, containerHeight, 5, 5, 'F');
    doc.setFontSize(11);
    doc.setFont(undefined, 'normal');
    doc.text('Predikat', 60, yPos + 25);
    doc.setFontSize(20);
    doc.setFont(undefined, 'bold');
    doc.text(summaryData.predikat, 60, yPos + 50);

    // Container 2: Total Karyawan
    doc.setFillColor(230, 230, 240);
    doc.roundedRect(
      40 + containerWidth + 20,
      yPos,
      containerWidth,
      containerHeight,
      5,
      5,
      'F'
    );
    doc.setFontSize(11);
    doc.setFont(undefined, 'normal');
    doc.text('Total Guru & Karyawan', 60 + containerWidth + 20, yPos + 25);
    doc.setFontSize(18);
    doc.setFont(undefined, 'bold');
    doc.text(
      `${summaryData.totalKaryawan} orang`,
      60 + containerWidth + 20,
      yPos + 50
    );

    // Container 3: Total Hari Kerja
    doc.setFillColor(230, 230, 240);
    doc.roundedRect(
      40 + (containerWidth + 20) * 2,
      yPos,
      containerWidth,
      containerHeight,
      5,
      5,
      'F'
    );
    doc.setFontSize(11);
    doc.setFont(undefined, 'normal');
    doc.text('Total Hari Kerja', 60 + (containerWidth + 20) * 2, yPos + 25);
    doc.setFontSize(18);
    doc.setFont(undefined, 'bold');
    doc.text(
      `${summaryData.totalHariKerja} hari`,
      60 + (containerWidth + 20) * 2,
      yPos + 50
    );

    yPos += containerHeight + 20;

    // Tingkat Kehadiran (tambahan info)
    doc.setFillColor(220, 220, 240);
    doc.roundedRect(40, yPos, pageWidth - 80, 30, 5, 5, 'F');
    doc.setFontSize(14);
    doc.setFont(undefined, 'bold');
    doc.text(
      `Tingkat Kehadiran: ${summaryData.persentaseKehadiran}%`,
      60,
      yPos + 20
    );

    yPos += 50;

    // Rincian Kehadiran Header
    doc.setFontSize(14);
    doc.setFont(undefined, 'bold');
    doc.setFillColor(220, 220, 240);
    doc.roundedRect(40, yPos, pageWidth - 80, 30, 5, 5, 'F');
    doc.text('Rincian Kehadiran:', 60, yPos + 20);

    yPos += 40;

    const details = [
      {
        label: 'Hadir Total',
        value: `${summaryData.totalHadir} (${summaryData.persentaseKehadiran}%)`,
        color: [200, 230, 255],
      },
      {
        label: 'Tepat Waktu',
        value: `${summaryData.totalTepat} (${summaryData.persentaseTepat}%)`,
        color: [200, 255, 200],
      },
      {
        label: 'Terlambat',
        value: `${summaryData.totalTelat} (${summaryData.persentaseTelat}%)`,
        color: [255, 255, 200],
      },
      {
        label: 'Alfa',
        value: `${summaryData.totalAlfa} (${summaryData.persentaseAlfa}%)`,
        color: [255, 200, 200],
      },
    ];

    // 4 box dengan ukuran sama dan simetris
    const detailBoxWidth = (pageWidth - 120) / 4;
    const detailBoxHeight = 60;

    details.forEach((detail, idx) => {
      const xPos = 40 + idx * (detailBoxWidth + 10);
      doc.setFillColor(...detail.color);
      doc.roundedRect(xPos, yPos, detailBoxWidth, detailBoxHeight, 5, 5, 'F');

      doc.setFontSize(11);
      doc.setFont(undefined, 'normal');
      doc.text(detail.label, xPos + 15, yPos + 25);

      doc.setFontSize(14);
      doc.setFont(undefined, 'bold');
      doc.text(detail.value, xPos + 15, yPos + 45);
    });

    yPos += detailBoxHeight + 20;

    // Naikkan posisi container 0.5 cm (≈ 14 pt)
    yPos -= 8;

    // ============= DESKRIPSI KESIMPULAN DENGAN CONTAINER =============
    doc.setFillColor(245, 245, 250);
    doc.setDrawColor(200, 200, 220);
    doc.setLineWidth(1.5);
    doc.roundedRect(40, yPos, pageWidth - 80, 200, 5, 5, 'FD');

    // Judul
    yPos += 20;
    doc.setFontSize(12);
    doc.setFont(undefined, 'bold');
    doc.setTextColor(79, 70, 229);
    doc.text('Deskripsi Kesimpulan Profil Absensi', 60, yPos);

    // Isi awal deskripsi
    yPos += 20;
    doc.setTextColor(0, 0, 0);
    doc.setFont(undefined, 'normal');
    doc.setFontSize(9);

    // Paragraf Pembuka
    const introText = `Profil absensi periode ${summaryData.periode} menunjukkan tingkat kehadiran ${summaryData.persentaseKehadiran}% yang termasuk kategori ${summaryData.predikat}.`;
    const introLines = doc.splitTextToSize(introText, pageWidth - 120);
    doc.text(introLines, 60, yPos);
    yPos += introLines.length * 12 + 10;

    // Analisis Kehadiran
    doc.setFont(undefined, 'bold');
    doc.text('Analisis Kehadiran:', 60, yPos);
    doc.setFont(undefined, 'normal');
    yPos += 12;
    let analisisText = '';
    if (summaryData.persentaseKehadiran >= 96) {
      analisisText =
        'Tingkat kehadiran sangat luar biasa dengan konsistensi kehadiran hampir sempurna.';
    } else if (summaryData.persentaseKehadiran >= 91) {
      analisisText =
        'Tingkat kehadiran sangat memuaskan dengan komitmen tinggi dari guru & karyawan.';
    } else if (summaryData.persentaseKehadiran >= 86) {
      analisisText =
        'Tingkat kehadiran baik dan menunjukkan dedikasi yang konsisten.';
    } else if (summaryData.persentaseKehadiran >= 81) {
      analisisText =
        'Tingkat kehadiran cukup baik namun masih ada ruang perbaikan.';
    } else if (summaryData.persentaseKehadiran >= 76) {
      analisisText =
        'Tingkat kehadiran di bawah standar dengan cukup banyak ketidakhadiran.';
    } else {
      analisisText =
        'Tingkat kehadiran di bawah standar minimal dengan banyak ketidakhadiran tanpa keterangan jelas.';
    }
    const analisisLines = doc.splitTextToSize(analisisText, pageWidth - 120);
    doc.text(analisisLines, 60, yPos);
    yPos += analisisLines.length * 12 + 8;

    // Kesadaran Absensi
    doc.setFont(undefined, 'bold');
    doc.text('Kesadaran Absensi:', 60, yPos);
    doc.setFont(undefined, 'normal');
    yPos += 12;
    let kesadaranText = '';
    if (summaryData.persentaseKehadiran >= 96) {
      kesadaranText =
        'Ketertiban scan masuk-pulang sangat sempurna. Hampir semua guru & karyawan konsisten melakukan scan lengkap.';
    } else if (summaryData.persentaseKehadiran >= 91) {
      kesadaranText =
        'Ketertiban scan masuk-pulang sangat baik. Mayoritas konsisten melakukan scan lengkap setiap hari.';
    } else if (summaryData.persentaseKehadiran >= 86) {
      kesadaranText =
        'Ketertiban scan masuk-pulang baik. Sebagian besar melakukan scan dengan tertib.';
    } else if (summaryData.persentaseKehadiran >= 81) {
      kesadaranText =
        'Ketertiban perlu ditingkatkan. Masih ditemukan kasus lupa scan pulang atau tidak scan sama sekali.';
    } else if (summaryData.persentaseKehadiran >= 76) {
      kesadaranText =
        'Ketertiban scan masuk-pulang kurang. Cukup banyak guru & karyawan lupa scan pulang sehingga data tidak lengkap.';
    } else {
      kesadaranText =
        'Ketertiban scan masuk-pulang rendah. Banyak guru & karyawan lupa scan pulang sehingga data tidak lengkap, menunjukkan kurangnya kesadaran administrasi.';
    }
    const kesadaranLines = doc.splitTextToSize(kesadaranText, pageWidth - 120);
    doc.text(kesadaranLines, 60, yPos);
    yPos += kesadaranLines.length * 12 + 8;

    // Kedisiplinan Waktu
    doc.setFont(undefined, 'bold');
    doc.text('Kedisiplinan Waktu:', 60, yPos);
    doc.setFont(undefined, 'normal');
    yPos += 12;
    let kedisiplinanText = '';
    if (summaryData.persentaseKehadiran >= 96) {
      kedisiplinanText =
        'Ketepatan waktu sangat sempurna. Hampir semua datang sebelum jadwal yang ditentukan.';
    } else if (summaryData.persentaseKehadiran >= 91) {
      kedisiplinanText =
        'Ketepatan waktu sangat baik. Sebagian besar datang sebelum atau tepat jadwal yang ditentukan.';
    } else if (summaryData.persentaseKehadiran >= 86) {
      kedisiplinanText =
        'Ketepatan waktu baik. Mayoritas guru & karyawan datang tepat waktu sesuai jadwal.';
    } else if (summaryData.persentaseKehadiran >= 81) {
      kedisiplinanText =
        'Ketepatan waktu bervariasi. Sebagian disiplin namun masih ada yang sering terlambat.';
    } else if (summaryData.persentaseKehadiran >= 76) {
      kedisiplinanText =
        'Ketepatan waktu kurang. Cukup banyak guru & karyawan datang terlambat dari jadwal.';
    } else {
      kedisiplinanText =
        'Ketepatan waktu rendah. Banyak guru & karyawan datang terlambat setelah jadwal dimulai.';
    }
    const kedisiplinanLines = doc.splitTextToSize(
      kedisiplinanText,
      pageWidth - 120
    );
    doc.text(kedisiplinanLines, 60, yPos);
    yPos += kedisiplinanLines.length * 12 + 8;

    // Rekomendasi/Apresiasi
    doc.setFont(undefined, 'bold');
    const rekomendasiLabel =
      summaryData.persentaseKehadiran >= 91 ? 'Apresiasi:' : 'Rekomendasi:';
    doc.text(rekomendasiLabel, 60, yPos);
    doc.setFont(undefined, 'normal');
    yPos += 12;
    let rekomendasiText = '';
    if (summaryData.persentaseKehadiran >= 96) {
      rekomendasiText =
        'Prestasi luar biasa! Pertahankan kedisiplinan sempurna ini dan jadilah teladan bagi yang lain.';
    } else if (summaryData.persentaseKehadiran >= 91) {
      rekomendasiText =
        'Prestasi sangat baik! Pertahankan disiplin ini dan tingkatkan menuju level UNGGUL.';
    } else if (summaryData.persentaseKehadiran >= 86) {
      rekomendasiText =
        'Performa baik, pertahankan dan tingkatkan konsistensi untuk mencapai kategori BAIK SEKALI.';
    } else if (summaryData.persentaseKehadiran >= 81) {
      rekomendasiText =
        'Disarankan reminder rutin dan evaluasi berkala untuk mencapai kategori BAIK atau BAIK SEKALI.';
    } else if (summaryData.persentaseKehadiran >= 76) {
      rekomendasiText =
        'Perlu pembinaan dan monitoring ketat untuk perbaikan kedisiplinan secara bertahap.';
    } else {
      rekomendasiText =
        'Perlu evaluasi individual, pembinaan intensif, dan penerapan sanksi tegas untuk perbaikan kedisiplinan.';
    }
    const rekomendasiLines = doc.splitTextToSize(
      rekomendasiText,
      pageWidth - 120
    );
    doc.text(rekomendasiLines, 60, yPos);

    // ============= HALAMAN 2: RANKING (3 KOLOM) =============
    if (rankingData) {
      doc.addPage();
      yPos = 40;

      // Header Peringkat
      doc.setFillColor(79, 70, 229);
      doc.rect(0, 0, pageWidth, 60, 'F');
      doc.setTextColor(255, 255, 255);
      doc.setFontSize(22);
      doc.setFont(undefined, 'bold');
      doc.text('PERINGKAT GURU & KARYAWAN', 40, 25);

      doc.setFontSize(14);
      doc.setFont(undefined, 'bold');
      doc.text('MTs. AN-NUR BULULAWANG', 40, 40);
      doc.setFontSize(10);
      doc.setFont(undefined, 'normal');
      doc.text(`Periode: ${summaryData.periode}`, 40, 55);

      yPos = 80;

      // Lebar kolom untuk 3 tabel
      const colWidth = (pageWidth - 100) / 3;
      const startX = [40, 40 + colWidth + 10, 40 + 2 * (colWidth + 10)];

      // Tabel 1: Disiplin Waktu
      doc.setTextColor(255, 255, 255);
      doc.setFillColor(34, 197, 94); // Green
      doc.roundedRect(startX[0], yPos, colWidth, 25, 5, 5, 'F');
      doc.setFontSize(10);
      doc.setFont(undefined, 'bold');
      doc.text('Disiplin Waktu Tertinggi', startX[0] + 10, yPos + 17);

      // Tabel 2: Tertib Administrasi
      doc.setFillColor(59, 130, 246); // Blue
      doc.roundedRect(startX[1], yPos, colWidth, 25, 5, 5, 'F');
      doc.text('Tertib Administrasi', startX[1] + 10, yPos + 17);

      // Tabel 3: Rendah Kesadaran
      doc.setFillColor(239, 68, 68); // Red
      doc.roundedRect(startX[2], yPos, colWidth, 25, 5, 5, 'F');
      doc.text('Rendah Kesadaran', startX[2] + 10, yPos + 17);

      yPos += 50;

      // Data Ranking (max 10 baris)
      const maxRows = 10;
      doc.setTextColor(0, 0, 0);
      doc.setFontSize(8);
      const rowHeight = 30; // Tinggi baris 30px

      for (let i = 0; i < maxRows; i++) {
        const emp1 = rankingData.topDisiplin[i];
        const emp2 = rankingData.topTertib[i];
        const emp3 = rankingData.topRendah[i];

        // Kolom 1 - Disiplin Waktu
        if (emp1) {
          doc.setFillColor(i % 2 === 0 ? 240 : 255, 255, 240);
          doc.rect(startX[0], yPos, colWidth, rowHeight, 'F');
          doc.setFont(undefined, 'bold');
          doc.setFontSize(10);
          doc.text(`${i + 1}.`, startX[0] + 5, yPos + 17);
          doc.setFont(undefined, 'normal');
          doc.setFontSize(9);
          // Nama - potong jika terlalu panjang dan tambahkan ellipsis
          const maxNameLength = Math.floor(colWidth / 7);
          const displayName =
            emp1.name.length > maxNameLength
              ? emp1.name.substring(0, maxNameLength - 3) + '...'
              : emp1.name;
          doc.text(displayName, startX[0] + 24, yPos + 12);
          doc.setFontSize(7.5);
          doc.setTextColor(100, 100, 100);
          const maxPosLength = Math.floor(colWidth / 5.5);
          const displayPos =
            emp1.position.length > maxPosLength
              ? emp1.position.substring(0, maxPosLength - 3) + '...'
              : emp1.position;
          doc.text(displayPos, startX[0] + 24, yPos + 21);
          doc.setTextColor(0, 0, 0);
          doc.setFont(undefined, 'bold');
          doc.setFontSize(11);
          doc.text(
            `${emp1.persenHijau}%`,
            startX[0] + colWidth - 35,
            yPos + 17
          );
        }

        // Kolom 2 - Tertib Administrasi
        if (emp2) {
          doc.setFillColor(240, 248, i % 2 === 0 ? 255 : 250);
          doc.rect(startX[1], yPos, colWidth, rowHeight, 'F');
          doc.setFont(undefined, 'bold');
          doc.setFontSize(10);
          doc.text(`${i + 1}.`, startX[1] + 5, yPos + 17);
          doc.setFont(undefined, 'normal');
          doc.setFontSize(9);
          const maxNameLength = Math.floor(colWidth / 7);
          const displayName =
            emp2.name.length > maxNameLength
              ? emp2.name.substring(0, maxNameLength - 3) + '...'
              : emp2.name;
          doc.text(displayName, startX[1] + 24, yPos + 12);
          doc.setFontSize(7.5);
          doc.setTextColor(100, 100, 100);
          const maxPosLength = Math.floor(colWidth / 5.5);
          const displayPos =
            emp2.position.length > maxPosLength
              ? emp2.position.substring(0, maxPosLength - 3) + '...'
              : emp2.position;
          doc.text(displayPos, startX[1] + 24, yPos + 21);
          doc.setTextColor(0, 0, 0);
          doc.setFont(undefined, 'bold');
          doc.setFontSize(11);
          doc.text(`${emp2.persenBiru}%`, startX[1] + colWidth - 35, yPos + 17);
        }

        // Kolom 3 - Rendah Kesadaran
        if (emp3) {
          doc.setFillColor(255, i % 2 === 0 ? 240 : 250, 240);
          doc.rect(startX[2], yPos, colWidth, rowHeight, 'F');
          doc.setFont(undefined, 'bold');
          doc.setFontSize(10);
          doc.text(`${i + 1}.`, startX[2] + 5, yPos + 17);
          doc.setFont(undefined, 'normal');
          doc.setFontSize(9);
          const maxNameLength = Math.floor(colWidth / 7);
          const displayName =
            emp3.name.length > maxNameLength
              ? emp3.name.substring(0, maxNameLength - 3) + '...'
              : emp3.name;
          doc.text(displayName, startX[2] + 24, yPos + 12);
          doc.setFontSize(7.5);
          doc.setTextColor(100, 100, 100);
          const maxPosLength = Math.floor(colWidth / 5.5);
          const displayPos =
            emp3.position.length > maxPosLength
              ? emp3.position.substring(0, maxPosLength - 3) + '...'
              : emp3.position;
          doc.text(displayPos, startX[2] + 24, yPos + 21);
          doc.setTextColor(0, 0, 0);
          doc.setFont(undefined, 'bold');
          doc.setFontSize(11);
          doc.text(
            `${emp3.persenMerah}%`,
            startX[2] + colWidth - 35,
            yPos + 17
          );
        }

        yPos += rowHeight;
      }
    }

    // ============= HALAMAN 3-5: TABEL 1, 2, 3 =============
    for (let i = 1; i <= 3; i++) {
      doc.addPage();
      const tableId = `tabel${i}`;
      const table = document.getElementById(tableId);
      if (!table) continue;

      let tableTitle = '';
      if (i === 1) tableTitle = '1. REKAP MESIN (DATA MENTAH)';
      else if (i === 2) tableTitle = '2. KEDISIPLINAN WAKTU';
      else if (i === 3) tableTitle = '3. EVALUASI KEHADIRAN';

      // Capture Tabel
      const canvas = await html2canvas(table, {
        scale: 1.5, // Skala tabel tetap
        backgroundColor: '#ffffff',
        logging: false,
      });
      const imgData = canvas.toDataURL('image/jpeg', 0.85);
      const imgWidth = canvas.width;
      const imgHeight = canvas.height;

      // Header Halaman
      doc.setFillColor(79, 70, 229);
      doc.rect(0, 0, pageWidth, 60, 'F');
      doc.setTextColor(255, 255, 255);
      doc.setFontSize(20);
      doc.setFont(undefined, 'bold');
      doc.text('Sistem Rekap Absensi', 40, 25);
      doc.setFontSize(12);
      doc.setFont(undefined, 'normal');
      doc.text('MTs. AN-NUR BULULAWANG', 40, 45);

      doc.setTextColor(0, 0, 0);
      doc.setFontSize(14);
      doc.setFont(undefined, 'bold');
      doc.text(tableTitle, 40, 80);

      if (startDate && endDate) {
        doc.setFontSize(10);
        doc.setFont(undefined, 'normal');
        const periode = `Periode: ${new Date(startDate).toLocaleDateString(
          'id-ID'
        )} - ${new Date(endDate).toLocaleDateString('id-ID')}`;
        doc.text(periode, pageWidth - 220, 80);
      }

      const marginTop = 100;
      const marginLeft = 30;
      const marginRight = 30;
      const marginBottom = 50;

      // LOGIKA KHUSUS HALAMAN TABEL 3
      // Kita batasi lebar area tabel agar ada sisa ruang 320px di kanan untuk panduan (diperbesar)
      let effectivePageWidth = pageWidth;
      if (i === 3) {
        effectivePageWidth = pageWidth - 320;
      }

      const contentWidth = effectivePageWidth - marginLeft - marginRight;
      const contentHeight = pageHeight - marginTop - marginBottom;

      const ratio = Math.min(
        contentWidth / imgWidth,
        contentHeight / imgHeight
      );
      const scaledWidth = imgWidth * ratio;
      const scaledHeight = imgHeight * ratio;

      let yPosition = marginTop;
      doc.addImage(
        imgData,
        'JPEG',
        marginLeft,
        yPosition,
        scaledWidth,
        scaledHeight
      );

      // Logika untuk handling tabel panjang (paging)
      let heightLeft = scaledHeight - contentHeight;
      while (heightLeft > 0) {
        doc.addPage();
        // ... (Header halaman lanjutan sama seperti sebelumnya) ...
        doc.setFillColor(79, 70, 229);
        doc.rect(0, 0, pageWidth, 50, 'F');
        doc.setTextColor(255, 255, 255);
        doc.setFontSize(16);
        doc.setFont(undefined, 'bold');
        doc.text('MTs. AN-NUR BULULAWANG', 40, 20);
        doc.setFontSize(11);
        doc.setFont(undefined, 'normal');
        doc.text(`${tableTitle} (lanjutan)`, 40, 38);

        const nextPageMarginTop = 70;
        yPosition = nextPageMarginTop - (scaledHeight - heightLeft);
        doc.addImage(
          imgData,
          'JPEG',
          marginLeft,
          yPosition,
          scaledWidth,
          scaledHeight
        );
        heightLeft -= pageHeight - nextPageMarginTop - marginBottom;
      }
    }

    // ============= PANDUAN DI SEBELAH KANAN TABEL 3 =============
    // Kembali ke halaman terakhir (tempat Tabel 3 berada)
    const lastTablePage = doc.internal.getNumberOfPages();
    doc.setPage(lastTablePage);

    const panduanElement = panduanRef.current;
    if (panduanElement) {
      const panduanCanvas = await html2canvas(panduanElement, {
        scale: 2, // Scale tinggi agar gambar tajam
        backgroundColor: '#ffffff',
        logging: false,
      });
      const panduanImg = panduanCanvas.toDataURL('image/jpeg', 0.9);
      const panduanOriWidth = panduanCanvas.width;
      const panduanOriHeight = panduanCanvas.height;

      // --- PENGATURAN UKURAN PANDUAN ---
      // Target width diperbesar menjadi 280 (sebelumnya 200) agar jauh lebih jelas
      const targetPanduanWidth = 280;

      // Hitung rasio
      const pRatio = targetPanduanWidth / panduanOriWidth;

      const pScaledWidth = panduanOriWidth * pRatio;
      const pScaledHeight = panduanOriHeight * pRatio;

      // --- PENGATURAN POSISI PANDUAN ---
      // Posisikan mepet kanan dengan margin 20
      const pStartX = pageWidth - pScaledWidth - 20;
      const pStartY = 100;

      // --- RENDER ---
      doc.addImage(
        panduanImg,
        'JPEG',
        pStartX,
        pStartY,
        pScaledWidth,
        pScaledHeight
      );
    }

    // Footer semua halaman
    const totalPages = doc.internal.getNumberOfPages();
    for (let i = 1; i <= totalPages; i++) {
      doc.setPage(i);
      doc.setFontSize(9);
      doc.setTextColor(100, 100, 100);
      doc.text(`Halaman ${i} dari ${totalPages}`, 40, pageHeight - 20);
      doc.text(
        'Generated by Matsanuba Management Technology',
        pageWidth - 280,
        pageHeight - 20
      );
    }

    const timestamp = new Date().toISOString().split('T')[0];
    doc.save(`laporan_lengkap_absensi_${timestamp}.pdf`);
  };

  const downloadAsPdf = async (tableId, fileName, includeSummary = false) => {
    const table = document.getElementById(tableId);
    if (!table) {
      alert('Tabel tidak ditemukan!');
      return;
    }

    // Tentukan judul berdasarkan tableId
    let tableTitle = '';
    if (tableId === 'tabel1') {
      tableTitle = '1. REKAP MESIN (DATA MENTAH)';
    } else if (tableId === 'tabel2') {
      tableTitle = '2. KEDISIPLINAN WAKTU';
    } else if (tableId === 'tabel3') {
      tableTitle = '3. EVALUASI KEHADIRAN';
    }

    // Capture tabel sebagai canvas dengan html2canvas
    const canvas = await html2canvas(table, {
      scale: 1.8, // Quality lebih tinggi untuk 2-3 MB
      backgroundColor: '#ffffff',
      logging: false,
    });

    const imgData = canvas.toDataURL('image/jpeg', 0.9); // JPEG dengan quality 90%
    const imgWidth = canvas.width;
    const imgHeight = canvas.height;

    // Buat PDF landscape A4
    const doc = new jsPDF('l', 'px', 'a4');
    const pdfWidth = doc.internal.pageSize.getWidth();
    const pdfHeight = doc.internal.pageSize.getHeight();

    // Margin
    const marginLeft = 30;
    const marginTop = 100; // Perbesar margin atas untuk header
    const marginRight = 30;
    const marginBottom = 50; // Perbesar margin bawah untuk footer

    // Area konten
    const contentWidth = pdfWidth - marginLeft - marginRight;
    const contentHeight = pdfHeight - marginTop - marginBottom;

    // Header - Logo dan Judul (hanya di halaman pertama)
    doc.setFillColor(79, 70, 229); // Indigo
    doc.rect(0, 0, pdfWidth, 70, 'F');

    doc.setTextColor(255, 255, 255);
    doc.setFontSize(22);
    doc.setFont(undefined, 'bold');
    doc.text('Sistem Rekap Absensi', marginLeft, 30);

    doc.setFontSize(14);
    doc.setFont(undefined, 'normal');
    doc.text('MTs. AN-NUR BULULAWANG', marginLeft, 50);

    // Judul Tabel (di bawah header biru)
    doc.setTextColor(0, 0, 0);
    doc.setFontSize(16);
    doc.setFont(undefined, 'bold');
    doc.text(tableTitle, marginLeft, 85);

    // Periode (pojok kanan)
    if (startDate && endDate) {
      doc.setFontSize(11);
      doc.setFont(undefined, 'normal');
      const startD = new Date(startDate);
      const endD = new Date(endDate);
      const periode = `Periode: ${startD.toLocaleDateString(
        'id-ID'
      )} - ${endD.toLocaleDateString('id-ID')}`;
      doc.text(periode, pdfWidth - marginRight - 180, 85);
    }

    // Hitung skala agar fit di area konten
    const ratio = Math.min(contentWidth / imgWidth, contentHeight / imgHeight);
    const scaledWidth = imgWidth * ratio;
    const scaledHeight = imgHeight * ratio;

    // Tambahkan gambar ke PDF dengan margin
    let yPosition = marginTop;
    doc.addImage(
      imgData,
      'JPEG',
      marginLeft,
      yPosition,
      scaledWidth,
      scaledHeight
    );

    // Jika gambar lebih tinggi dari 1 halaman, buat multiple pages
    let heightLeft = scaledHeight - contentHeight;

    while (heightLeft > 0) {
      doc.addPage();

      // Header mini di halaman berikutnya
      doc.setFillColor(79, 70, 229);
      doc.rect(0, 0, pdfWidth, 60, 'F');
      doc.setTextColor(255, 255, 255);
      doc.setFontSize(18);
      doc.setFont(undefined, 'bold');
      doc.text('MTs. AN-NUR BULULAWANG', marginLeft, 25);
      doc.setFontSize(12);
      doc.setFont(undefined, 'normal');
      doc.text(`${tableTitle} (lanjutan)`, marginLeft, 45);

      // Margin atas untuk halaman lanjutan
      const nextPageMarginTop = 80;
      yPosition = nextPageMarginTop - (scaledHeight - heightLeft);
      doc.addImage(
        imgData,
        'JPEG',
        marginLeft,
        yPosition,
        scaledWidth,
        scaledHeight
      );
      heightLeft -= pdfHeight - nextPageMarginTop - marginBottom;
    }

    // Footer di semua halaman
    const pageCount = doc.internal.getNumberOfPages();
    for (let i = 1; i <= pageCount; i++) {
      doc.setPage(i);
      doc.setFontSize(9);
      doc.setTextColor(100, 100, 100);
      doc.text(`Halaman ${i} dari ${pageCount}`, marginLeft, pdfHeight - 20);
      doc.text(
        'Generated by Matsanuba Management Technology',
        pdfWidth - marginRight - 250,
        pdfHeight - 20
      );
    }

    doc.save(`${fileName}.pdf`);
  };

  const rgbToHex = (rgb) => {
    // Handle null/undefined
    if (!rgb) return null;

    // Handle jika rgb adalah object (dari className match)
    if (typeof rgb === 'object') return null;

    // Handle string
    const rgbString = String(rgb);

    // Jika sudah hex format
    if (rgbString.startsWith('#')) return rgbString;

    // Parse RGB format: rgb(255, 255, 255)
    const match = rgbString.match(/\d+/g);
    if (match && match.length >= 3) {
      return (
        '#' +
        match
          .slice(0, 3)
          .map((x) => parseInt(x).toString(16).padStart(2, '0'))
          .join('')
      );
    }

    return null;
  };

  const downloadSummaryAsPdf = () => {
    if (!summaryData) return;

    const doc = new jsPDF('p', 'pt', 'a4');
    const pageWidth = doc.internal.pageSize.getWidth();
    let yPos = 50;

    // Header dengan warna gradient (simulasi)
    doc.setFillColor(79, 70, 229); // Indigo
    doc.rect(0, 0, pageWidth, 100, 'F');

    doc.setTextColor(255, 255, 255);
    doc.setFontSize(22);
    doc.setFont(undefined, 'bold');
    // Hilangkan emoji, gunakan text saja
    doc.text('KESIMPULAN PROFIL ABSENSI', 40, yPos);

    doc.setFontSize(14);
    doc.setFont(undefined, 'normal');
    doc.text(`MTs. An-Nur Bululawang`, 40, yPos + 25);

    yPos = 120;
    doc.setTextColor(0, 0, 0);

    // Periode dan Predikat
    doc.setFillColor(240, 240, 255);
    doc.roundedRect(40, yPos, pageWidth - 80, 100, 5, 5, 'F');

    yPos += 25;
    doc.setFontSize(14);
    doc.setFont(undefined, 'bold');
    doc.text(`Periode: ${summaryData.periode}`, 60, yPos);

    yPos += 30;
    doc.setFontSize(20);
    doc.text(summaryData.predikat, 60, yPos);

    yPos += 30;
    doc.setFontSize(16);
    doc.text(
      `Tingkat Kehadiran: ${summaryData.persentaseKehadiran}%`,
      60,
      yPos
    );

    yPos += 50;

    // Total Karyawan dan Hari Kerja
    doc.setFontSize(12);
    doc.setFont(undefined, 'normal');

    doc.setFillColor(230, 230, 240);
    doc.roundedRect(40, yPos, (pageWidth - 100) / 2, 60, 5, 5, 'F');
    doc.roundedRect(
      40 + (pageWidth - 100) / 2 + 20,
      yPos,
      (pageWidth - 100) / 2,
      60,
      5,
      5,
      'F'
    );

    doc.setFont(undefined, 'normal');
    doc.text('Total Karyawan', 60, yPos + 25);
    doc.setFont(undefined, 'bold');
    doc.setFontSize(18);
    doc.text(`${summaryData.totalKaryawan} orang`, 60, yPos + 45);

    doc.setFontSize(12);
    doc.setFont(undefined, 'normal');
    doc.text('Total Hari Kerja', 60 + (pageWidth - 100) / 2 + 20, yPos + 25);
    doc.setFont(undefined, 'bold');
    doc.setFontSize(18);
    doc.text(
      `${summaryData.totalHariKerja} hari`,
      60 + (pageWidth - 100) / 2 + 20,
      yPos + 45
    );

    yPos += 80;

    // Rincian Kehadiran
    doc.setFontSize(14);
    doc.setFont(undefined, 'bold');
    doc.setFillColor(220, 220, 240);
    doc.roundedRect(40, yPos, pageWidth - 80, 30, 5, 5, 'F');
    doc.text('Rincian Kehadiran:', 60, yPos + 20);

    yPos += 50;

    const details = [
      {
        label: 'Hadir Total',
        value: `${summaryData.totalHadir} (${summaryData.persentaseKehadiran}%)`,
        color: [200, 230, 255],
      },
      {
        label: 'Tepat Waktu',
        value: `${summaryData.totalTepat} (${summaryData.persentaseTepat}%)`,
        color: [200, 255, 200],
      },
      {
        label: 'Terlambat',
        value: `${summaryData.totalTelat} (${summaryData.persentaseTelat}%)`,
        color: [255, 255, 200],
      },
      {
        label: 'Alfa',
        value: `${summaryData.totalAlfa} (${summaryData.persentaseAlfa}%)`,
        color: [255, 200, 200],
      },
    ];

    // Gunakan ukuran yang sama dengan box Total Karyawan & Total Hari Kerja
    const boxWidth = (pageWidth - 100) / 2;
    const boxHeight = 60;
    let col = 0;
    let row = 0;

    details.forEach((detail, index) => {
      const xPos = 40 + col * (boxWidth + 20);
      const yBoxPos = yPos + row * (boxHeight + 10);

      doc.setFillColor(...detail.color);
      doc.roundedRect(xPos, yBoxPos, boxWidth, boxHeight, 5, 5, 'F');

      doc.setFontSize(11);
      doc.setFont(undefined, 'normal');
      doc.text(detail.label, xPos + 15, yBoxPos + 25);

      doc.setFontSize(14);
      doc.setFont(undefined, 'bold');
      doc.text(detail.value, xPos + 15, yBoxPos + 45);

      col++;
      if (col >= 2) {
        col = 0;
        row++;
      }
    });

    yPos += 140;

    // Deskripsi Kesimpulan Profil Absensi dengan Container
    doc.setFillColor(245, 245, 250);
    doc.setDrawColor(200, 200, 220);
    doc.setLineWidth(1.5);
    doc.roundedRect(40, yPos, pageWidth - 80, 240, 5, 5, 'FD');

    yPos += 20;
    doc.setFontSize(12);
    doc.setFont(undefined, 'bold');
    doc.setTextColor(79, 70, 229);
    doc.text('Deskripsi Kesimpulan Profil Absensi', 60, yPos);

    yPos += 20;
    doc.setTextColor(0, 0, 0);
    doc.setFont(undefined, 'normal');
    doc.setFontSize(9);

    // Paragraf Pembuka
    const interpretasiText = `Profil absensi periode ${summaryData.periode} menunjukkan tingkat kehadiran ${summaryData.persentaseKehadiran}% yang termasuk kategori ${summaryData.predikat}.`;
    const interpretasiLines = doc.splitTextToSize(
      interpretasiText,
      pageWidth - 120
    );
    doc.text(interpretasiLines, 60, yPos);
    yPos += interpretasiLines.length * 12 + 10;

    // Analisis Kehadiran
    doc.setFont(undefined, 'bold');
    doc.text('Analisis Kehadiran:', 60, yPos);
    doc.setFont(undefined, 'normal');
    yPos += 12;
    let analisisText = '';
    if (summaryData.persentaseKehadiran >= 96) {
      analisisText =
        'Tingkat kehadiran sangat luar biasa dengan konsistensi kehadiran hampir sempurna.';
    } else if (summaryData.persentaseKehadiran >= 91) {
      analisisText =
        'Tingkat kehadiran sangat memuaskan dengan komitmen tinggi dari guru & karyawan.';
    } else if (summaryData.persentaseKehadiran >= 86) {
      analisisText =
        'Tingkat kehadiran baik dan menunjukkan dedikasi yang konsisten.';
    } else if (summaryData.persentaseKehadiran >= 81) {
      analisisText =
        'Tingkat kehadiran cukup baik namun masih ada ruang perbaikan.';
    } else if (summaryData.persentaseKehadiran >= 76) {
      analisisText =
        'Tingkat kehadiran di bawah standar dengan cukup banyak ketidakhadiran.';
    } else {
      analisisText =
        'Tingkat kehadiran di bawah standar minimal dengan banyak ketidakhadiran tanpa keterangan jelas.';
    }
    const analisisLines = doc.splitTextToSize(analisisText, pageWidth - 120);
    doc.text(analisisLines, 60, yPos);
    yPos += analisisLines.length * 12 + 8;

    // Kesadaran Absensi
    doc.setFont(undefined, 'bold');
    doc.text('Kesadaran Absensi:', 60, yPos);
    doc.setFont(undefined, 'normal');
    yPos += 12;
    let kesadaranText = '';
    if (summaryData.persentaseKehadiran >= 96) {
      kesadaranText =
        'Ketertiban scan masuk-pulang sangat sempurna. Hampir semua guru & karyawan konsisten melakukan scan lengkap.';
    } else if (summaryData.persentaseKehadiran >= 91) {
      kesadaranText =
        'Ketertiban scan masuk-pulang sangat baik. Mayoritas konsisten melakukan scan lengkap setiap hari.';
    } else if (summaryData.persentaseKehadiran >= 86) {
      kesadaranText =
        'Ketertiban scan masuk-pulang baik. Sebagian besar melakukan scan dengan tertib.';
    } else if (summaryData.persentaseKehadiran >= 81) {
      kesadaranText =
        'Ketertiban perlu ditingkatkan. Masih ditemukan kasus lupa scan pulang atau tidak scan sama sekali.';
    } else if (summaryData.persentaseKehadiran >= 76) {
      kesadaranText =
        'Ketertiban scan masuk-pulang kurang. Cukup banyak guru & karyawan lupa scan pulang sehingga data tidak lengkap.';
    } else {
      kesadaranText =
        'Ketertiban scan masuk-pulang rendah. Banyak guru & karyawan lupa scan pulang sehingga data tidak lengkap, menunjukkan kurangnya kesadaran administrasi.';
    }
    const kesadaranLines = doc.splitTextToSize(kesadaranText, pageWidth - 120);
    doc.text(kesadaranLines, 60, yPos);
    yPos += kesadaranLines.length * 12 + 8;

    // Kedisiplinan Waktu
    doc.setFont(undefined, 'bold');
    doc.text('Kedisiplinan Waktu:', 60, yPos);
    doc.setFont(undefined, 'normal');
    yPos += 12;
    let kedisiplinanText = '';
    if (summaryData.persentaseKehadiran >= 96) {
      kedisiplinanText =
        'Ketepatan waktu sangat sempurna. Hampir semua datang sebelum jadwal yang ditentukan.';
    } else if (summaryData.persentaseKehadiran >= 91) {
      kedisiplinanText =
        'Ketepatan waktu sangat baik. Sebagian besar datang sebelum atau tepat jadwal yang ditentukan.';
    } else if (summaryData.persentaseKehadiran >= 86) {
      kedisiplinanText =
        'Ketepatan waktu baik. Mayoritas guru & karyawan datang tepat waktu sesuai jadwal.';
    } else if (summaryData.persentaseKehadiran >= 81) {
      kedisiplinanText =
        'Ketepatan waktu bervariasi. Sebagian disiplin namun masih ada yang sering terlambat.';
    } else if (summaryData.persentaseKehadiran >= 76) {
      kedisiplinanText =
        'Ketepatan waktu kurang. Cukup banyak guru & karyawan datang terlambat dari jadwal.';
    } else {
      kedisiplinanText =
        'Ketepatan waktu rendah. Banyak guru & karyawan datang terlambat setelah jadwal dimulai.';
    }
    const kedisiplinanLines = doc.splitTextToSize(
      kedisiplinanText,
      pageWidth - 120
    );
    doc.text(kedisiplinanLines, 60, yPos);
    yPos += kedisiplinanLines.length * 12 + 8;

    // Rekomendasi/Apresiasi
    doc.setFont(undefined, 'bold');
    const rekomendasiLabel =
      summaryData.persentaseKehadiran >= 91 ? 'Apresiasi:' : 'Rekomendasi:';
    doc.text(rekomendasiLabel, 60, yPos);
    doc.setFont(undefined, 'normal');
    yPos += 12;
    let rekomendasiText = '';
    if (summaryData.persentaseKehadiran >= 96) {
      rekomendasiText =
        'Prestasi luar biasa! Pertahankan kedisiplinan sempurna ini dan jadilah teladan bagi yang lain.';
    } else if (summaryData.persentaseKehadiran >= 91) {
      rekomendasiText =
        'Prestasi sangat baik! Pertahankan disiplin ini dan tingkatkan menuju level UNGGUL.';
    } else if (summaryData.persentaseKehadiran >= 86) {
      rekomendasiText =
        'Performa baik, pertahankan dan tingkatkan konsistensi untuk mencapai kategori BAIK SEKALI.';
    } else if (summaryData.persentaseKehadiran >= 81) {
      rekomendasiText =
        'Disarankan reminder rutin dan evaluasi berkala untuk mencapai kategori BAIK atau BAIK SEKALI.';
    } else if (summaryData.persentaseKehadiran >= 76) {
      rekomendasiText =
        'Perlu pembinaan dan monitoring ketat untuk perbaikan kedisiplinan secara bertahap.';
    } else {
      rekomendasiText =
        'Perlu evaluasi individual, pembinaan intensif, dan penerapan sanksi tegas untuk perbaikan kedisiplinan.';
    }
    const rekomendasiLines = doc.splitTextToSize(
      rekomendasiText,
      pageWidth - 120
    );
    doc.text(rekomendasiLines, 60, yPos);

    // Tambahkan 3 Tabel Ranking jika ada
    if (rankingData) {
      // SELALU pindah ke halaman baru untuk peringkat
      doc.addPage();
      yPos = 40;

      // Header Peringkat (tanpa emoji)
      doc.setFillColor(79, 70, 229);
      doc.roundedRect(40, yPos, pageWidth - 80, 40, 5, 5, 'F');
      doc.setTextColor(255, 255, 255);
      doc.setFontSize(16);
      doc.setFont(undefined, 'bold');
      doc.text('PERINGKAT GURU & KARYAWAN', 60, yPos + 25);

      yPos += 60;

      // Tabel 1: Disiplin Waktu Tertinggi
      doc.setTextColor(0, 0, 0);
      doc.setFontSize(12);
      doc.setFont(undefined, 'bold');
      doc.setFillColor(34, 197, 94); // Green
      doc.roundedRect(40, yPos, pageWidth - 80, 25, 5, 5, 'F');
      doc.setTextColor(255, 255, 255);
      doc.text(
        '1. Peringkat Disiplin Waktu Tertinggi (Datang Sebelum Jadwal)',
        60,
        yPos + 17
      );

      yPos += 35;

      const table1Data = rankingData.topDisiplin.map((emp, idx) => [
        idx + 1,
        emp.id,
        emp.name,
        emp.position,
        emp.hijau,
        emp.persenHijau + '%',
      ]);

      autoTable(doc, {
        head: [
          ['Peringkat', 'ID', 'Nama', 'Jabatan', 'Total Hijau', 'Persentase'],
        ],
        body: table1Data,
        startY: yPos,
        theme: 'grid',
        headStyles: {
          fillColor: [34, 197, 94],
          fontStyle: 'bold',
          fontSize: 8,
        },
        styles: { fontSize: 7, cellPadding: 2.5 },
        columnStyles: {
          0: { halign: 'center', cellWidth: 25 },
          1: { halign: 'center', cellWidth: 35 },
          2: { cellWidth: 'auto' },
          3: { cellWidth: 'auto' },
          4: { halign: 'center', cellWidth: 35 },
          5: { halign: 'center', cellWidth: 40 },
        },
      });

      yPos = doc.lastAutoTable.finalY + 15;

      // Cek halaman baru sebelum tabel 2
      if (yPos > 680) {
        doc.addPage();
        yPos = 40;
      }

      // Tabel 2: Tertib Administrasi
      doc.setFontSize(12);
      doc.setFont(undefined, 'bold');
      doc.setFillColor(59, 130, 246); // Blue
      doc.roundedRect(40, yPos, pageWidth - 80, 25, 5, 5, 'F');
      doc.setTextColor(255, 255, 255);
      doc.text(
        '2. Peringkat Tertib Administrasi (Scan Masuk & Pulang Lengkap)',
        60,
        yPos + 17
      );

      yPos += 35;

      const table2Data = rankingData.topTertib.map((emp, idx) => [
        idx + 1,
        emp.id,
        emp.name,
        emp.position,
        emp.biru,
        emp.persenBiru + '%',
      ]);

      autoTable(doc, {
        head: [
          ['Peringkat', 'ID', 'Nama', 'Jabatan', 'Total Biru', 'Persentase'],
        ],
        body: table2Data,
        startY: yPos,
        theme: 'grid',
        headStyles: {
          fillColor: [59, 130, 246],
          fontStyle: 'bold',
          fontSize: 8,
        },
        styles: { fontSize: 7, cellPadding: 2.5 },
        columnStyles: {
          0: { halign: 'center', cellWidth: 25 },
          1: { halign: 'center', cellWidth: 35 },
          2: { cellWidth: 'auto' },
          3: { cellWidth: 'auto' },
          4: { halign: 'center', cellWidth: 35 },
          5: { halign: 'center', cellWidth: 40 },
        },
      });

      yPos = doc.lastAutoTable.finalY + 15;

      // Cek halaman baru sebelum tabel 3
      if (yPos > 680) {
        doc.addPage();
        yPos = 40;
      }

      // Tabel 3: Rendah Kesadaran
      doc.setFontSize(12);
      doc.setFont(undefined, 'bold');
      doc.setFillColor(239, 68, 68); // Red
      doc.roundedRect(40, yPos, pageWidth - 80, 25, 5, 5, 'F');
      doc.setTextColor(255, 255, 255);
      doc.text(
        '3. Peringkat Rendah Kesadaran Absensi (Alfa/Tidak Scan)',
        60,
        yPos + 17
      );

      yPos += 35;

      const table3Data = rankingData.topRendah.map((emp, idx) => [
        idx + 1,
        emp.id,
        emp.name,
        emp.position,
        emp.merah,
        emp.persenMerah + '%',
      ]);

      autoTable(doc, {
        head: [
          ['Peringkat', 'ID', 'Nama', 'Jabatan', 'Total Merah', 'Persentase'],
        ],
        body: table3Data,
        startY: yPos,
        theme: 'grid',
        headStyles: {
          fillColor: [239, 68, 68],
          fontStyle: 'bold',
          fontSize: 8,
        },
        styles: { fontSize: 7, cellPadding: 2.5 },
        columnStyles: {
          0: { halign: 'center', cellWidth: 25 },
          1: { halign: 'center', cellWidth: 35 },
          2: { cellWidth: 'auto' },
          3: { cellWidth: 'auto' },
          4: { halign: 'center', cellWidth: 35 },
          5: { halign: 'center', cellWidth: 40 },
        },
      });
    }

    doc.save('kesimpulan_profil_absensi.pdf');
  };
  const downloadAllTablesAsExcel = () => {
    if (!recapData) {
      alert('Silakan generate tabel rekap terlebih dahulu');
      return;
    }

    // Helper function untuk clone dan style table
    const prepareTableHTML = (tableId, title) => {
      const table = document.getElementById(tableId);
      if (!table) return '';

      const clonedTable = table.cloneNode(true);

      // Inject inline styles ke setiap cell
      const allCells = clonedTable.querySelectorAll('td, th');
      allCells.forEach((cell) => {
        const classList = cell.className;
        let bgColor = '#FFFFFF';
        let fontWeight = 'normal';
        let textAlign = 'center';

        // Deteksi warna dari className
        if (
          classList.includes('bg-blue-200') ||
          classList.includes('bg-blue-100')
        ) {
          bgColor = '#ADD8E6';
        } else if (classList.includes('bg-blue-300')) {
          bgColor = '#93C5FD';
        } else if (
          classList.includes('bg-yellow-200') ||
          classList.includes('bg-yellow-100')
        ) {
          bgColor = '#FFFF99';
        } else if (classList.includes('bg-yellow-300')) {
          bgColor = '#FDE047';
        } else if (
          classList.includes('bg-red-200') ||
          classList.includes('bg-red-100')
        ) {
          bgColor = '#FFB3B3';
        } else if (
          classList.includes('bg-green-200') ||
          classList.includes('bg-green-100')
        ) {
          bgColor = '#90EE90';
        } else if (classList.includes('bg-green-300')) {
          bgColor = '#86EFAC';
        } else if (classList.includes('bg-gray-100')) {
          bgColor = '#F3F4F6';
        } else if (classList.includes('bg-gray-600')) {
          bgColor = '#4B5563';
        } else if (classList.includes('bg-gray-700')) {
          bgColor = '#374151';
        } else if (classList.includes('bg-purple-100')) {
          bgColor = '#E9D5FF';
        } else if (classList.includes('bg-indigo-100')) {
          bgColor = '#C7D2FE';
        } else if (classList.includes('bg-gray-300')) {
          bgColor = '#D1D5DB';
        }

        // Check inline background color
        if (cell.style.backgroundColor) {
          const inlineColor = cell.style.backgroundColor;
          if (inlineColor.startsWith('rgb')) {
            const match = inlineColor.match(/\d+/g);
            if (match && match.length >= 3) {
              bgColor =
                '#' +
                match
                  .slice(0, 3)
                  .map((x) => parseInt(x).toString(16).padStart(2, '0'))
                  .join('');
            }
          } else if (inlineColor.startsWith('#')) {
            bgColor = inlineColor;
          }
        }

        // Deteksi bold
        if (classList.includes('font-bold')) {
          fontWeight = 'bold';
        }

        // Set inline style untuk Excel
        cell.setAttribute(
          'style',
          `background-color: ${bgColor}; font-weight: ${fontWeight}; text-align: ${textAlign}; border: 1px solid #000000; padding: 5px;`
        );
      });

      // Wrap dengan div dan title
      return `
        <div style="display: inline-block; vertical-align: top; margin-right: 30px;">
          <h3 style="background-color: #4F46E5; color: white; padding: 10px; text-align: center; margin: 0;">${title}</h3>
          <table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse;">
            ${clonedTable.innerHTML}
          </table>
        </div>
      `;
    };

    // Prepare 3 tabel dengan title
    const table1HTML = prepareTableHTML('tabel1', '1. REKAP MESIN');
    const table2HTML = prepareTableHTML('tabel2', '2. KEDISIPLINAN WAKTU');
    const table3HTML = prepareTableHTML('tabel3', '3. EVALUASI KEHADIRAN');

    // Gunakan template literal tanpa indentasi yang bikin JSX curiga
    let html = `
<html>
<head>
  <meta charset="utf-8">
  <style>
    body { font-family: Arial, sans-serif; }
    table { border-collapse: collapse; }
    td, th { border: 1px solid #000000; padding: 5px; }
  </style>
</head>
<body>
  <div style="white-space: nowrap; overflow-x: auto;">
    ${table1HTML}
    ${table2HTML}
    ${table3HTML}
  </div>
</body>
</html>
`;

    // Create blob dan download
    const blob = new Blob([html], {
      type: 'application/vnd.ms-excel;charset=utf-8',
    });

    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    const timestamp = new Date().toISOString().split('T')[0];
    link.download = `rekap_absensi_lengkap_${timestamp}.xls`;
    link.click();
    URL.revokeObjectURL(link.href);

    setTimeout(() => {
      alert(
        '✅ Excel berhasil didownload!\n\n💡 Tips: Scroll ke bawah untuk melihat tabel 2 dan 3'
      );
    }, 500);
  };

  const renderSummary = () => {
    if (!summaryData) return null;

    return (
      <div
        ref={summaryRef}
        className={
          'p-4 sm:p-6 rounded-lg sm:rounded-xl shadow-lg text-white bg-gradient-to-br mb-4 sm:mb-6 md:mb-8 ' +
          summaryData.warna
        }
      >
        <h3 className="text-lg sm:text-xl md:text-2xl font-bold mb-3 sm:mb-4 flex items-center gap-2">
          {summaryData.icon} KESIMPULAN PROFIL ABSENSI
        </h3>
        {/* MODIFIKASI DISINI: 
          Menggunakan Grid 3 Kolom untuk:
          1. Periode/Status
          2. Total Guru
          3. Hari Kerja
        */}
        <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3 sm:gap-4 mb-3 sm:mb-4">
          {/* KOLOM 1: Periode & Status */}
          <div className="bg-white bg-opacity-20 rounded-lg p-3 sm:p-4 md:p-5 flex flex-col justify-center">
            <p className="text-sm sm:text-base md:text-lg font-semibold mb-1 opacity-90">
              Periode: {summaryData.periode}
            </p>
            <p className="text-4xl font-bold my-2">{summaryData.predikat}</p>
            <p className="text-lg opacity-90">
              Tingkat Kehadiran: {summaryData.persentaseKehadiran}%
            </p>
          </div>

          {/* KOLOM 2: Total Guru */}
          <div className="bg-white bg-opacity-20 rounded-lg p-5 flex flex-col justify-center">
            <p className="text-xs sm:text-sm opacity-90 mb-1">
              Total Guru & Karyawan
            </p>
            <p className="text-2xl sm:text-3xl font-bold">
              {summaryData.totalKaryawan} orang
            </p>
          </div>

          {/* KOLOM 3: Hari Kerja */}
          <div className="bg-white bg-opacity-20 rounded-lg p-5 flex flex-col justify-center">
            <p className="text-sm opacity-90 mb-1">
              Hari Kerja Guru & Karyawan
            </p>
            <p className="text-3xl font-bold">
              {summaryData.totalHariKerja} hari
            </p>
          </div>
        </div>

        {/* Bagian Bawah: Rincian Kehadiran (Tetap full width) */}
        <div className="bg-white bg-opacity-20 rounded-lg p-3 sm:p-4 md:p-5">
          <p className="font-semibold mb-3 sm:mb-4 text-base sm:text-lg">
            Rincian Kehadiran:
          </p>
          <div className="grid grid-cols-2 md:grid-cols-4 gap-3 sm:gap-4 text-xs sm:text-sm">
            <div>
              <p className="opacity-80 text-xs uppercase tracking-wider mb-1">
                Hadir Total
              </p>
              <p className="text-xl font-bold">
                {summaryData.totalHadir}{' '}
                <span className="text-base font-normal opacity-90">
                  ({summaryData.persentaseKehadiran}%)
                </span>
              </p>
            </div>
            <div>
              <p className="opacity-80 text-xs uppercase tracking-wider mb-1">
                Tepat Waktu
              </p>
              <p className="text-xl font-bold">
                {summaryData.totalTepat}{' '}
                <span className="text-base font-normal opacity-90">
                  ({summaryData.persentaseTepat}%)
                </span>
              </p>
            </div>
            <div>
              <p className="opacity-80 text-xs uppercase tracking-wider mb-1">
                Terlambat
              </p>
              <p className="text-xl font-bold">
                {summaryData.totalTelat}{' '}
                <span className="text-base font-normal opacity-90">
                  ({summaryData.persentaseTelat}%)
                </span>
              </p>
            </div>
            <div>
              <p className="opacity-80 text-xs uppercase tracking-wider mb-1">
                Alfa
              </p>
              <p className="text-xl font-bold">
                {summaryData.totalAlfa}{' '}
                <span className="text-base font-normal opacity-90">
                  ({summaryData.persentaseAlfa}%)
                </span>
              </p>
            </div>
          </div>
        </div>

        <div className="mt-3 sm:mt-4 p-2.5 sm:p-3 bg-white bg-opacity-10 rounded-lg text-xs sm:text-sm">
          <p className="opacity-90 mb-2">
            <strong>Deskripsi Kesimpulan Profil Absensi</strong>
          </p>
          <p className="opacity-90 mb-4">
            Profil absensi periode {summaryData.periode} menunjukkan tingkat
            kehadiran {summaryData.persentaseKehadiran}% yang termasuk kategori{' '}
            <strong>{summaryData.predikat}</strong>.
          </p>

          {/* Layout 3 Kolom dengan Grid */}
          <div className="space-y-3">
            {/* Analisis Kehadiran */}
            <div className="grid grid-cols-[160px_auto_1fr] gap-2 opacity-90">
              <div className="font-bold">Analisis Kehadiran</div>
              <div>:</div>
              <div className="flex-1">
                {summaryData.persentaseKehadiran >= 96 &&
                  'Tingkat kehadiran sangat luar biasa dengan konsistensi kehadiran hampir sempurna.'}
                {summaryData.persentaseKehadiran >= 91 &&
                  summaryData.persentaseKehadiran < 96 &&
                  'Tingkat kehadiran sangat memuaskan dengan komitmen tinggi dari guru & karyawan.'}
                {summaryData.persentaseKehadiran >= 86 &&
                  summaryData.persentaseKehadiran < 91 &&
                  'Tingkat kehadiran baik dan menunjukkan dedikasi yang konsisten.'}
                {summaryData.persentaseKehadiran >= 81 &&
                  summaryData.persentaseKehadiran < 86 &&
                  'Tingkat kehadiran cukup baik namun masih ada ruang perbaikan.'}
                {summaryData.persentaseKehadiran >= 76 &&
                  summaryData.persentaseKehadiran < 81 &&
                  'Tingkat kehadiran di bawah standar dengan cukup banyak ketidakhadiran.'}
                {summaryData.persentaseKehadiran < 76 &&
                  'Tingkat kehadiran di bawah standar minimal dengan banyak ketidakhadiran tanpa keterangan jelas.'}
              </div>
            </div>

            {/* Kesadaran Absensi */}
            <div className="grid grid-cols-[160px_auto_1fr] gap-2 opacity-90">
              <div className="font-bold">Kesadaran Absensi</div>
              <div>:</div>{' '}
              <div className="flex-1">
                {summaryData.persentaseKehadiran >= 96 &&
                  'Ketertiban scan masuk-pulang sangat sempurna. Hampir semua guru & karyawan konsisten melakukan scan lengkap.'}
                {summaryData.persentaseKehadiran >= 91 &&
                  summaryData.persentaseKehadiran < 96 &&
                  'Ketertiban scan masuk-pulang sangat baik. Mayoritas konsisten melakukan scan lengkap setiap hari.'}
                {summaryData.persentaseKehadiran >= 86 &&
                  summaryData.persentaseKehadiran < 91 &&
                  'Ketertiban scan masuk-pulang baik. Sebagian besar melakukan scan dengan tertib.'}
                {summaryData.persentaseKehadiran >= 81 &&
                  summaryData.persentaseKehadiran < 86 &&
                  'Ketertiban perlu ditingkatkan. Masih ditemukan kasus lupa scan pulang atau tidak scan sama sekali.'}
                {summaryData.persentaseKehadiran >= 76 &&
                  summaryData.persentaseKehadiran < 81 &&
                  'Ketertiban scan masuk-pulang kurang. Cukup banyak guru & karyawan lupa scan pulang sehingga data tidak lengkap.'}
                {summaryData.persentaseKehadiran < 76 &&
                  'Ketertiban scan masuk-pulang rendah. Banyak guru & karyawan lupa scan pulang sehingga data tidak lengkap, menunjukkan kurangnya kesadaran administrasi.'}
              </div>
            </div>

            {/* Kedisiplinan Waktu */}
            <div className="grid grid-cols-[160px_auto_1fr] gap-2 opacity-90">
              <div className="font-bold">Kedisiplinan Waktu</div>
              <div>:</div>
              <div className="flex-1">
                {summaryData.persentaseKehadiran >= 96 &&
                  'Ketepatan waktu sangat sempurna. Hampir semua datang sebelum jadwal yang ditentukan.'}
                {summaryData.persentaseKehadiran >= 91 &&
                  summaryData.persentaseKehadiran < 96 &&
                  'Ketepatan waktu sangat baik. Sebagian besar datang sebelum atau tepat jadwal yang ditentukan.'}
                {summaryData.persentaseKehadiran >= 86 &&
                  summaryData.persentaseKehadiran < 91 &&
                  'Ketepatan waktu baik. Mayoritas guru & karyawan datang tepat waktu sesuai jadwal.'}
                {summaryData.persentaseKehadiran >= 81 &&
                  summaryData.persentaseKehadiran < 86 &&
                  'Ketepatan waktu bervariasi. Sebagian disiplin namun masih ada yang sering terlambat.'}
                {summaryData.persentaseKehadiran >= 76 &&
                  summaryData.persentaseKehadiran < 81 &&
                  'Ketepatan waktu kurang. Cukup banyak guru & karyawan datang terlambat dari jadwal.'}
                {summaryData.persentaseKehadiran < 76 &&
                  'Ketepatan waktu rendah. Banyak guru & karyawan datang terlambat setelah jadwal dimulai.'}
              </div>
            </div>

            {/* Rekomendasi/Apresiasi */}
            <div className="grid grid-cols-[160px_auto_1fr] gap-2 opacity-90">
              <div className="font-bold">
                {summaryData.persentaseKehadiran >= 91
                  ? 'Apresiasi'
                  : 'Rekomendasi'}
              </div>
              <div>:</div>
              <div className="flex-1">
                {summaryData.persentaseKehadiran >= 96 &&
                  'Prestasi luar biasa! Pertahankan kedisiplinan sempurna ini dan jadilah teladan bagi yang lain.'}
                {summaryData.persentaseKehadiran >= 91 &&
                  summaryData.persentaseKehadiran < 96 &&
                  'Prestasi sangat baik! Pertahankan disiplin ini dan tingkatkan menuju level UNGGUL.'}
                {summaryData.persentaseKehadiran >= 86 &&
                  summaryData.persentaseKehadiran < 91 &&
                  'Performa baik, pertahankan dan tingkatkan konsistensi untuk mencapai kategori BAIK SEKALI.'}
                {summaryData.persentaseKehadiran >= 81 &&
                  summaryData.persentaseKehadiran < 86 &&
                  'Disarankan reminder rutin dan evaluasi berkala untuk mencapai kategori BAIK atau BAIK SEKALI.'}
                {summaryData.persentaseKehadiran >= 76 &&
                  summaryData.persentaseKehadiran < 81 &&
                  'Perlu pembinaan dan monitoring ketat untuk perbaikan kedisiplinan secara bertahap.'}
                {summaryData.persentaseKehadiran < 76 &&
                  'Perlu evaluasi individual, pembinaan intensif, dan penerapan sanksi tegas untuk perbaikan kedisiplinan.'}
              </div>
            </div>
          </div>
        </div>

        {/* 3 TABEL PERINGKAT */}
        {rankingData && (
          <div className="mt-4 sm:mt-6 space-y-4 sm:space-y-6">
            <h3 className="text-lg sm:text-xl font-bold text-white mb-3 sm:mb-4">
              🏆 PERINGKAT KARYAWAN
            </h3>

            {/* Tabel 1: Disiplin Waktu Tertinggi */}
            <div className="bg-white bg-opacity-20 rounded-lg p-3 sm:p-4">
              <h4 className="font-bold text-white text-sm sm:text-base mb-2 sm:mb-3 flex items-center gap-2">
                <span className="bg-green-500 text-white px-3 py-1 rounded-full text-sm">
                  1
                </span>
                Peringkat Disiplin Waktu Tertinggi (Datang Sebelum Jadwal)
              </h4>
              <div className="overflow-x-auto -mx-3 sm:mx-0">
                <div className="min-w-max px-3 sm:px-0">
                  <table className="w-full text-xs sm:text-sm bg-white rounded-lg">
                    <thead className="bg-green-600 text-white">
                      <tr>
                        <th className="px-2 sm:px-3 py-1.5 sm:py-2 text-center w-16 sm:w-20 text-xs sm:text-sm">
                          Peringkat
                        </th>
                        <th className="px-3 py-2 text-center w-24">ID</th>
                        <th className="px-3 py-2 text-left w-64">Nama</th>
                        <th className="px-3 py-2 text-left w-48">Jabatan</th>
                        <th className="px-3 py-2 text-center w-28">
                          Total Hijau
                        </th>
                        <th className="px-3 py-2 text-center w-28">
                          Persentase
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {rankingData.topDisiplin.map((emp, idx) => (
                        <tr
                          key={emp.id}
                          className={idx % 2 === 0 ? 'bg-green-50' : 'bg-white'}
                        >
                          <td className="px-2 sm:px-3 py-1.5 sm:py-2 text-center font-bold text-green-700 text-xs sm:text-sm">
                            {idx + 1}
                          </td>
                          <td className="px-3 py-2 text-center text-gray-700">
                            {emp.id}
                          </td>
                          <td className="px-3 py-2 text-gray-800">
                            {emp.name}
                          </td>
                          <td className="px-3 py-2 text-gray-600 text-sm">
                            {emp.position}
                          </td>
                          <td className="px-3 py-2 text-center font-bold text-green-600">
                            {emp.hijau}
                          </td>
                          <td className="px-3 py-2 text-center font-bold text-green-600">
                            {emp.persenHijau}%
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>

            {/* Tabel 2: Tertib Administrasi */}
            <div className="bg-white bg-opacity-20 rounded-lg p-3 sm:p-4">
              <h4 className="font-bold text-white text-sm sm:text-base mb-2 sm:mb-3 flex items-center gap-2">
                <span className="bg-blue-500 text-white px-3 py-1 rounded-full text-sm">
                  2
                </span>
                Peringkat Tertib Administrasi (Scan Masuk & Pulang Lengkap)
              </h4>
              <div className="overflow-x-auto -mx-3 sm:mx-0">
                <div className="min-w-max px-3 sm:px-0">
                  <table className="w-full text-xs sm:text-sm bg-white rounded-lg">
                    <thead className="bg-blue-600 text-white">
                      <tr>
                        <th className="px-2 sm:px-3 py-1.5 sm:py-2 text-center w-16 sm:w-20 text-xs sm:text-sm">
                          Peringkat
                        </th>
                        <th className="px-3 py-2 text-center w-24">ID</th>
                        <th className="px-3 py-2 text-left w-64">Nama</th>
                        <th className="px-3 py-2 text-left w-48">Jabatan</th>
                        <th className="px-3 py-2 text-center w-28">
                          Total Biru
                        </th>
                        <th className="px-3 py-2 text-center w-28">
                          Persentase
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {rankingData.topTertib.map((emp, idx) => (
                        <tr
                          key={emp.id}
                          className={idx % 2 === 0 ? 'bg-blue-50' : 'bg-white'}
                        >
                          <td className="px-3 py-2 text-center font-bold text-blue-700">
                            {idx + 1}
                          </td>
                          <td className="px-3 py-2 text-center text-gray-700">
                            {emp.id}
                          </td>
                          <td className="px-3 py-2 text-gray-800">
                            {emp.name}
                          </td>
                          <td className="px-3 py-2 text-gray-600 text-sm">
                            {emp.position}
                          </td>
                          <td className="px-3 py-2 text-center font-bold text-blue-600">
                            {emp.biru}
                          </td>
                          <td className="px-3 py-2 text-center font-bold text-blue-600">
                            {emp.persenBiru}%
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>

            {/* Tabel 3: Rendah Kesadaran */}
            <div className="bg-white bg-opacity-20 rounded-lg p-3 sm:p-4">
              <h4 className="font-bold text-white text-sm sm:text-base mb-2 sm:mb-3 flex items-center gap-2">
                <span className="bg-red-500 text-white px-3 py-1 rounded-full text-sm">
                  3
                </span>
                Peringkat Rendah Kesadaran Absensi (Alfa/Tidak Scan)
              </h4>
              <div className="overflow-x-auto -mx-3 sm:mx-0">
                <div className="min-w-max px-3 sm:px-0">
                  <table className="w-full text-xs sm:text-sm bg-white rounded-lg">
                    <thead className="bg-red-600 text-white">
                      <tr>
                        <th className="px-2 sm:px-3 py-1.5 sm:py-2 text-center w-16 sm:w-20 text-xs sm:text-sm">
                          Peringkat
                        </th>
                        <th className="px-3 py-2 text-center w-24">ID</th>
                        <th className="px-3 py-2 text-left w-64">Nama</th>
                        <th className="px-3 py-2 text-left w-48">Jabatan</th>
                        <th className="px-3 py-2 text-center w-28">
                          Total Merah
                        </th>
                        <th className="px-3 py-2 text-center w-28">
                          Persentase
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {rankingData.topRendah.map((emp, idx) => (
                        <tr
                          key={emp.id}
                          className={idx % 2 === 0 ? 'bg-red-50' : 'bg-white'}
                        >
                          <td className="px-3 py-2 text-center font-bold text-red-700">
                            {idx + 1}
                          </td>
                          <td className="px-3 py-2 text-center text-gray-700">
                            {emp.id}
                          </td>
                          <td className="px-3 py-2 text-gray-800">
                            {emp.name}
                          </td>
                          <td className="px-3 py-2 text-gray-600 text-sm">
                            {emp.position}
                          </td>
                          <td className="px-3 py-2 text-center font-bold text-red-600">
                            {emp.merah}
                          </td>
                          <td className="px-3 py-2 text-center font-bold text-red-600">
                            {emp.persenMerah}%
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          </div>
        )}

        <div className="mt-3 sm:mt-4 flex flex-wrap gap-2 sm:gap-3">
          <button
            onClick={downloadSummaryAsPdf}
            className="bg-white bg-opacity-20 text-white text-xs sm:text-sm px-3 sm:px-4 py-2 rounded-lg hover:bg-opacity-30 flex items-center gap-1.5 sm:gap-2 flex-1 sm:flex-initial justify-center"
          >
            <Download size={16} /> PDF
          </button>
          <button
            onClick={downloadAllTablesAsExcel}
            className="bg-white bg-opacity-20 text-white px-4 py-2 rounded-lg hover:bg-opacity-30 flex items-center gap-2"
          >
            <Download size={16} /> All Table
          </button>
          <button
            onClick={handleCopySummary}
            className="bg-white bg-opacity-20 text-white px-4 py-2 rounded-lg hover:bg-opacity-30 flex items-center gap-2"
          >
            <FileText size={16} /> Copy JPG
          </button>
          <button
            onClick={handleDownloadSummaryJPG}
            className="bg-white bg-opacity-20 text-white px-4 py-2 rounded-lg hover:bg-opacity-30 flex items-center gap-2"
          >
            <Download size={16} /> Download JPG
          </button>
        </div>
      </div>
    );
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-100 p-3 sm:p-6 md:p-8">
      <div className="max-w-7xl mx-auto">
        <div className="bg-white rounded-xl sm:rounded-2xl shadow-xl p-4 sm:p-6 md:p-8">
          <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between mb-4 sm:mb-6 gap-4">
            <div className="flex items-center gap-3 sm:gap-4 w-full sm:w-auto">
              <img
                src="https://numufidz.github.io/overlay/logo_mtsannur.png"
                alt="Logo MTs. An-Nur Bululawang"
                className="w-12 h-12 sm:w-16 sm:h-16 object-contain flex-shrink-0"
              />
              <div className="flex-1 min-w-0">
                <h1 className="text-xl sm:text-2xl md:text-3xl font-bold text-indigo-800 truncate">
                  Sistem Rekap Absensi
                </h1>
                <h2 className="text-base sm:text-lg md:text-xl font-semibold text-gray-800 truncate">
                  MTs. An-Nur Bululawang
                </h2>
                <p className="text-xs sm:text-sm text-gray-600 truncate">
                  Jl. Diponegoro IV Bululawang
                </p>
              </div>
            </div>
            <div className="text-left sm:text-right text-xs sm:text-sm text-gray-500 w-full sm:w-auto">
              <p>Powered by:</p>
              <p>Matsanuba Management Technology</p>
              <p>Version 1.0 | © 2025</p>
            </div>
          </div>
          <p className="text-sm sm:text-base text-gray-600 mb-4 sm:mb-6 md:mb-8 border-t pt-3 sm:pt-4">
            Sistem profesional untuk evaluasi absensi karyawan berdasarkan data
            mesin fingerprint dan jadwal kerja. Menyediakan analisis
            kedisiplinan, rekap mentah, dan peringkat performa secara akurat dan
            efisien.
          </p>
          {errorMessage && <p className="text-red-600 mb-4">{errorMessage}</p>}

          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 sm:gap-6 mb-4 sm:mb-6 md:mb-8">
            <div className="border-2 border-dashed border-indigo-300 rounded-lg sm:rounded-xl p-4 sm:p-6 relative">
              <label className="cursor-pointer block">
                <div className="flex flex-col items-center gap-2 sm:gap-3">
                  <Upload className="text-indigo-600" size={32} />
                  <span className="text-base sm:text-lg font-semibold text-gray-700 text-center">
                    Laporan Absensi
                  </span>
                </div>
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={(e) =>
                    e.target.files &&
                    e.target.files[0] &&
                    processAttendanceFile(e.target.files[0])
                  }
                  className="hidden"
                />
              </label>
              {isLoadingAttendance && (
                <div className="absolute inset-0 flex items-center justify-center bg-white bg-opacity-50">
                  <Loader className="animate-spin text-indigo-600" size={40} />
                </div>
              )}
              {attendanceData.length > 0 && (
                <div className="mt-3 sm:mt-4 p-2 sm:p-3 bg-green-50 border border-green-200 rounded-lg">
                  <p className="text-green-800 font-medium text-xs sm:text-sm">
                    Data dimuat
                  </p>
                </div>
              )}
            </div>
            <div className="border-2 border-dashed border-purple-300 rounded-xl p-6 relative">
              <label className="cursor-pointer block">
                <div className="flex flex-col items-center gap-3">
                  <Upload className="text-purple-600" size={40} />
                  <span className="text-lg font-semibold text-gray-700">
                    Jadwal Kerja
                  </span>
                </div>
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={(e) =>
                    e.target.files &&
                    e.target.files[0] &&
                    processScheduleFile(e.target.files[0])
                  }
                  className="hidden"
                />
              </label>
              {isLoadingSchedule && (
                <div className="absolute inset-0 flex items-center justify-center bg-white bg-opacity-50">
                  <Loader className="animate-spin text-purple-600" size={40} />
                </div>
              )}
              {scheduleData.length > 0 && (
                <div className="mt-4 p-3 bg-green-50 border border-green-200 rounded-lg">
                  <p className="text-green-800 font-medium text-sm">
                    {scheduleData.length} jadwal
                  </p>
                </div>
              )}
            </div>
          </div>
          <div className="bg-indigo-50 rounded-lg sm:rounded-xl p-4 sm:p-6 mb-4 sm:mb-6 md:mb-8">
            <div className="mb-6">
              <h4 className="font-medium text-gray-700 mb-3">Periode</h4>
              <div className="flex flex-col sm:flex-row gap-3 sm:gap-4">
                <div className="flex-1">
                  <label className="block text-xs sm:text-sm font-medium text-gray-700 mb-1 sm:mb-2">
                    Tanggal Awal
                  </label>
                  <input
                    type="date"
                    value={startDate}
                    onChange={(e) => setStartDate(e.target.value)}
                    className="w-full px-4 py-2 border border-gray-300 rounded-lg"
                  />
                </div>
                <div className="flex-1">
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    Tanggal Akhir
                  </label>
                  <input
                    type="date"
                    value={endDate}
                    onChange={(e) => setEndDate(e.target.value)}
                    className="w-full px-4 py-2 border border-gray-300 rounded-lg"
                  />
                </div>
              </div>
            </div>
            <div>
              <h4 className="font-medium text-gray-700 mb-3">Generate</h4>
              <div className="flex flex-col sm:flex-row gap-2 sm:gap-4">
                {/* 1. Tabel Rekap */}
                <button
                  onClick={generateRecapTables}
                  disabled={attendanceData.length === 0}
                  className="flex-1 bg-indigo-600 text-white font-semibold text-sm sm:text-base 
               py-2.5 sm:py-3 px-4 sm:px-6 rounded-lg sm:rounded-xl 
               hover:bg-indigo-700 disabled:bg-gray-300 transition-colors 
               flex items-center justify-center gap-2"
                >
                  <span className="text-base sm:text-lg">📊</span>
                  Tabel Rekap
                </button>

                {/* 2. Kesimpulan Profil */}
                <button
                  onClick={generateSummary}
                  disabled={!recapData}
                  className="flex-1 bg-purple-600 text-white font-semibold text-sm sm:text-base 
               py-2.5 sm:py-3 px-4 sm:px-6 rounded-lg sm:rounded-xl 
               hover:bg-purple-700 disabled:bg-gray-300 transition-colors 
               flex items-center justify-center gap-2"
                >
                  <span className="text-base sm:text-lg">🧠</span>
                  Kesimpulan Profil
                </button>
              </div>
              {/* 3. PDF Lengkap (tanpa simbol tambahan) */}
              <button
                onClick={downloadCompletePdf}
                disabled={!recapData}
                className="w-full mt-2 bg-green-600 text-white font-semibold text-sm sm:text-base
             py-3 sm:py-4 px-4 sm:px-6 rounded-lg sm:rounded-xl
             hover:bg-green-700 transition-colors shadow-md
             flex items-center justify-center gap-2
             disabled:bg-gray-300 disabled:text-white disabled:cursor-not-allowed"
              >
                <Download size={20} />
                PDF Lengkap
              </button>{' '}
            </div>
          </div>
          {summaryData && renderSummary()}
          {recapData && (
            <div className="mt-8 border-t pt-8 space-y-8">
              <div>
                <div className="bg-blue-100 p-3 sm:p-4 rounded-lg mb-3 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3">
                  <h3 className="text-lg sm:text-xl font-bold text-gray-800">
                    1. Rekap Mesin
                  </h3>
                  <div className="flex flex-wrap gap-2 w-full sm:w-auto">
                    <button
                      onClick={() => copyTableToClipboard('tabel1')}
                      className="bg-blue-600 text-white px-4 py-2 rounded-lg hover:bg-blue-700 flex items-center gap-2"
                    >
                      <FileText size={16} /> Copy
                    </button>
                    <button
                      onClick={() =>
                        downloadTableAsExcel('tabel1', 'rekap_mesin')
                      }
                      className="bg-blue-600 text-white px-4 py-2 rounded-lg hover:bg-blue-700 flex items-center gap-2"
                    >
                      <Download size={16} /> Excel
                    </button>
                    <button
                      onClick={() => downloadAsPdf('tabel1', 'rekap_mesin')}
                      className="bg-blue-600 text-white px-4 py-2 rounded-lg hover:bg-blue-700 flex items-center gap-2"
                    >
                      <Download size={16} /> PDF
                    </button>
                  </div>
                </div>
                <div className="overflow-x-auto border rounded-lg mb-4 sm:mb-6 md:mb-8 -mx-4 sm:mx-0">
                  <div className="min-w-max px-4 sm:px-0">
                    <table
                      id="tabel1"
                      className="min-w-full text-sm border-collapse bg-white"
                    >
                      <thead>
                        <tr className="bg-gray-300">
                          <th className="border border-gray-400 px-1.5 sm:px-2 py-1.5 sm:py-2 text-black font-bold text-xs sm:text-sm">
                            No
                          </th>
                          <th className="border border-gray-400 px-1.5 sm:px-2 py-1.5 sm:py-2 text-black font-bold text-xs sm:text-sm">
                            ID
                          </th>
                          <th
                            className="border border-gray-400 px-2 py-2 text-black font-bold"
                            style={{ minWidth: '220px' }}
                          >
                            Nama
                          </th>
                          <th className="border border-gray-400 px-6 py-3 text-black font-bold min-w-[120px]">
                            Jabatan
                          </th>
                          {recapData.dateRange.map((date) => (
                            <th
                              key={date}
                              className="border border-gray-400 px-2 py-2 text-black font-bold"
                              colSpan={2}
                            >
                              {getDateLabel(date)}
                            </th>
                          ))}
                          <th className="border border-gray-400 px-2 py-2 text-black font-bold bg-gray-300">
                            Hari Kerja
                          </th>
                          <th className="border border-gray-400 px-2 py-2 text-black font-bold bg-blue-200">
                            Biru
                          </th>
                          <th className="border border-gray-400 px-2 py-2 text-black font-bold bg-yellow-200">
                            Kuning
                          </th>
                          <th className="border border-gray-400 px-2 py-2 text-black font-bold bg-red-200">
                            Merah
                          </th>
                          <th className="border border-gray-400 px-2 py-2 text-black font-bold bg-indigo-200">
                            Hadir
                          </th>
                          <th className="border border-gray-400 px-2 py-2 text-black font-bold bg-purple-200">
                            %
                          </th>
                        </tr>
                        <tr className="bg-gray-100">
                          <th
                            className="border border-gray-400"
                            colSpan={4}
                          ></th>
                          {recapData.dateRange.map((date) => (
                            <React.Fragment key={date}>
                              <th className="border border-gray-400 px-1 py-1 text-gray-700 text-xs">
                                Masuk
                              </th>
                              <th className="border border-gray-400 px-1 py-1 text-gray-700 text-xs">
                                Pulang
                              </th>
                            </React.Fragment>
                          ))}
                          <th
                            className="border border-gray-400"
                            colSpan={6}
                          ></th>
                        </tr>{' '}
                      </thead>
                      <tbody>
                        {recapData.recap.map((emp) => {
                          let hariKerja = 0;
                          let biru = 0;
                          let kuning = 0;
                          let merah = 0;

                          recapData.dateRange.forEach((dateStr) => {
                            const ev = emp.dailyEvaluation[dateStr];
                            if (ev.text !== 'L') {
                              hariKerja++;
                              const rec = emp.dailyRecords[dateStr];
                              const hasIn = rec.in !== '-';
                              const hasOut = rec.out !== '-';

                              if (hasIn && hasOut) biru++;
                              else if (!hasIn && !hasOut) merah++;
                              else kuning++;
                            }
                          });

                          const hadir = biru + kuning;
                          const persentase =
                            hariKerja > 0
                              ? Math.round((hadir / hariKerja) * 100)
                              : 0;

                          return (
                            <tr key={emp.no}>
                              <td className="border border-gray-400 px-1.5 sm:px-2 py-1.5 sm:py-2 text-center text-xs sm:text-sm">
                                {emp.no}
                              </td>
                              <td className="border border-gray-400 px-1.5 sm:px-2 py-1.5 sm:py-2 text-center text-xs sm:text-sm">
                                {emp.id}
                              </td>
                              <td className="border border-gray-400 px-2 py-2">
                                {emp.name}
                              </td>
                              <td className="border border-gray-400 px-2 py-2">
                                {emp.position}
                              </td>
                              {recapData.dateRange.map((date) => {
                                const rec = emp.dailyRecords[date];
                                const ev = emp.dailyEvaluation[date];
                                if (ev.text === 'L') {
                                  return (
                                    <>
                                      <td
                                        key={date + '-in'}
                                        className="border border-gray-400 text-center font-bold"
                                      >
                                        L
                                      </td>
                                      <td
                                        key={date + '-out'}
                                        className="border border-gray-400 text-center font-bold"
                                      >
                                        L
                                      </td>
                                    </>
                                  );
                                }
                                const hasIn = rec.in !== '-';
                                const hasOut = rec.out !== '-';
                                const bg =
                                  hasIn && hasOut
                                    ? 'bg-blue-200'
                                    : !hasIn && !hasOut
                                    ? 'bg-red-200'
                                    : 'bg-yellow-200';
                                return (
                                  <React.Fragment key={date}>
                                    <td
                                      className={
                                        'border border-gray-400 px-1 py-1 text-center text-xs ' +
                                        bg
                                      }
                                    >
                                      {rec.in}
                                    </td>
                                    <td
                                      className={
                                        'border border-gray-400 px-1 py-1 text-center text-xs ' +
                                        bg
                                      }
                                    >
                                      {rec.out}
                                    </td>
                                  </React.Fragment>
                                );
                              })}
                              <td className="border border-gray-400 px-2 py-2 text-center font-bold bg-gray-100">
                                {hariKerja}
                              </td>
                              <td className="border border-gray-400 px-2 py-2 text-center font-bold bg-blue-100">
                                {biru}
                              </td>
                              <td className="border border-gray-400 px-2 py-2 text-center font-bold bg-yellow-100">
                                {kuning}
                              </td>
                              <td className="border border-gray-400 px-2 py-2 text-center font-bold bg-red-100">
                                {merah}
                              </td>
                              <td className="border border-gray-400 px-2 py-2 text-center font-bold bg-indigo-100">
                                {hadir}
                              </td>
                              <td className="border border-gray-400 px-2 py-2 text-center font-bold bg-purple-100">
                                {persentase}%
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
              <div>
                <div className="bg-green-100 p-3 sm:p-4 rounded-lg mb-3 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3">
                  <h3 className="text-lg sm:text-xl font-bold text-gray-800">
                    2. Kedisiplinan Waktu
                  </h3>
                  <div className="flex flex-wrap gap-2 w-full sm:w-auto">
                    <button
                      onClick={() => copyTableToClipboard('tabel2')}
                      className="bg-green-600 text-white px-4 py-2 rounded-lg hover:bg-green-700 flex items-center gap-2"
                    >
                      <FileText size={16} /> Copy
                    </button>
                    <button
                      onClick={() =>
                        downloadTableAsExcel('tabel2', 'kedisiplinan_waktu')
                      }
                      className="bg-green-600 text-white px-4 py-2 rounded-lg hover:bg-green-700 flex items-center gap-2"
                    >
                      <Download size={16} /> Excel
                    </button>
                    <button
                      onClick={() =>
                        downloadAsPdf('tabel2', 'kedisiplinan_waktu')
                      }
                      className="bg-green-600 text-white px-4 py-2 rounded-lg hover:bg-green-700 flex items-center gap-2"
                    >
                      <Download size={16} /> PDF
                    </button>
                  </div>
                </div>
                <div className="overflow-x-auto border rounded-lg mb-4 sm:mb-6 md:mb-8 -mx-4 sm:mx-0">
                  <div className="min-w-max px-4 sm:px-0">
                    <table
                      id="tabel2"
                      className="min-w-full text-sm border-collapse bg-white"
                    >
                      <thead>
                        <tr className="bg-gray-300">
                          <th className="border border-gray-400 px-1.5 sm:px-2 py-1.5 sm:py-2 text-black font-bold text-xs sm:text-sm">
                            No
                          </th>
                          <th className="border border-gray-400 px-1.5 sm:px-2 py-1.5 sm:py-2 text-black font-bold text-xs sm:text-sm">
                            ID
                          </th>
                          <th
                            className="border border-gray-400 px-2 py-2 text-black font-bold"
                            style={{ minWidth: '220px' }}
                          >
                            Nama
                          </th>
                          <th className="border border-gray-400 px-6 py-3 text-black font-bold min-w-[120px]">
                            Jabatan
                          </th>
                          {recapData.dateRange.map((date) => (
                            <th
                              key={date}
                              className="border border-gray-400 px-2 py-2 text-black font-bold"
                              colSpan={2}
                            >
                              {getDateLabel(date)}
                            </th>
                          ))}
                          <th className="border border-gray-400 px-2 py-2 text-black font-bold bg-gray-300">
                            Hari Kerja
                          </th>
                          <th className="border border-gray-400 px-2 py-2 text-black font-bold bg-green-200">
                            Hijau
                          </th>
                          <th className="border border-gray-400 px-2 py-2 text-black font-bold bg-blue-200">
                            Biru
                          </th>
                          <th className="border border-gray-400 px-2 py-2 text-black font-bold bg-yellow-200">
                            Kuning
                          </th>
                          <th className="border border-gray-400 px-2 py-2 text-black font-bold bg-red-200">
                            Merah
                          </th>
                          <th className="border border-gray-400 px-2 py-2 text-black font-bold bg-purple-200">
                            %
                          </th>
                        </tr>
                        <tr className="bg-gray-100">
                          <th
                            className="border border-gray-400"
                            colSpan={4}
                          ></th>
                          {recapData.dateRange.map((date) => (
                            <React.Fragment key={date}>
                              <th className="border border-gray-400 px-1 py-1 text-gray-700 text-xs">
                                Masuk
                              </th>
                              <th className="border border-gray-400 px-1 py-1 text-gray-700 text-xs">
                                Pulang
                              </th>
                            </React.Fragment>
                          ))}
                          <th
                            className="border border-gray-400"
                            colSpan={6}
                          ></th>
                        </tr>
                      </thead>
                      <tbody>
                        {recapData.recap.map((emp) => {
                          const sched = scheduleData.find(
                            (s) => s.id === emp.id
                          );
                          let hariKerja = 0;
                          let hijau = 0;
                          let biru = 0;
                          let kuning = 0;
                          let merah = 0;

                          recapData.dateRange.forEach((dateStr) => {
                            const ev = emp.dailyEvaluation[dateStr];
                            if (ev.text !== 'L') {
                              hariKerja++;
                              const rec = emp.dailyRecords[dateStr];
                              const hasIn = rec.in !== '-';
                              const hasOut = rec.out !== '-';

                              if (!hasIn && !hasOut) {
                                merah++;
                              } else if (hasIn && hasOut) {
                                biru++;
                              } else if (hasIn) {
                                const dayName = getDayName(dateStr);
                                const schedStart =
                                  sched?.schedule[dayName]?.start;
                                const inMin = timeToMinutes(rec.in);
                                const schedMin = timeToMinutes(schedStart);
                                if (schedMin && inMin && inMin <= schedMin) {
                                  hijau++;
                                } else {
                                  kuning++;
                                }
                              } else {
                                kuning++;
                              }
                            }
                          });

                          const disiplin = hijau + biru;
                          const persentase =
                            hariKerja > 0
                              ? Math.round((disiplin / hariKerja) * 100)
                              : 0;

                          return (
                            <tr key={emp.no}>
                              <td className="border border-gray-400 px-1.5 sm:px-2 py-1.5 sm:py-2 text-center text-xs sm:text-sm">
                                {emp.no}
                              </td>
                              <td className="border border-gray-400 px-1.5 sm:px-2 py-1.5 sm:py-2 text-center text-xs sm:text-sm">
                                {emp.id}
                              </td>
                              <td className="border border-gray-400 px-2 py-2">
                                {emp.name}
                              </td>
                              <td className="border border-gray-400 px-2 py-2">
                                {emp.position}
                              </td>
                              {recapData.dateRange.map((date) => {
                                const rec = emp.dailyRecords[date];
                                const ev = emp.dailyEvaluation[date];
                                if (ev.text === 'L') {
                                  return (
                                    <>
                                      <td
                                        key={date + '-in'}
                                        className="border border-gray-400 text-center font-bold"
                                      >
                                        L
                                      </td>
                                      <td
                                        key={date + '-out'}
                                        className="border border-gray-400 text-center font-bold"
                                      >
                                        L
                                      </td>
                                    </>
                                  );
                                }
                                const hasIn = rec.in !== '-';
                                const hasOut = rec.out !== '-';
                                let bg = 'bg-red-200';
                                if (hasIn && hasOut) {
                                  bg = 'bg-blue-300';
                                } else if (hasIn) {
                                  const dayName = getDayName(date);
                                  const schedStart =
                                    sched?.schedule[dayName]?.start;
                                  const inMin = timeToMinutes(rec.in);
                                  const schedMin = timeToMinutes(schedStart);
                                  if (schedMin && inMin && inMin <= schedMin) {
                                    bg = 'bg-green-300';
                                  } else {
                                    bg = 'bg-yellow-300';
                                  }
                                } else if (hasOut) {
                                  bg = 'bg-yellow-300';
                                }
                                return (
                                  <React.Fragment key={date}>
                                    <td
                                      className={
                                        'border border-gray-400 px-1 py-1 text-center text-xs ' +
                                        bg
                                      }
                                    >
                                      {rec.in}
                                    </td>
                                    <td
                                      className={
                                        'border border-gray-400 px-1 py-1 text-center text-xs ' +
                                        bg
                                      }
                                    >
                                      {rec.out}
                                    </td>
                                  </React.Fragment>
                                );
                              })}
                              <td className="border border-gray-400 px-2 py-2 text-center font-bold bg-gray-100">
                                {hariKerja}
                              </td>
                              <td className="border border-gray-400 px-2 py-2 text-center font-bold bg-green-100">
                                {hijau}
                              </td>
                              <td className="border border-gray-400 px-2 py-2 text-center font-bold bg-blue-100">
                                {biru}
                              </td>
                              <td className="border border-gray-400 px-2 py-2 text-center font-bold bg-yellow-100">
                                {kuning}
                              </td>
                              <td className="border border-gray-400 px-2 py-2 text-center font-bold bg-red-100">
                                {merah}
                              </td>
                              <td className="border border-gray-400 px-2 py-2 text-center font-bold bg-purple-100">
                                {persentase}%
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
              <div>
                <div className="bg-purple-100 p-3 sm:p-4 rounded-lg mb-3 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3">
                  <h3 className="text-lg sm:text-xl font-bold text-gray-800">
                    3. Evaluasi Kehadiran
                  </h3>
                  <div className="flex flex-wrap gap-2 w-full sm:w-auto">
                    <button
                      onClick={() => copyTableToClipboard('tabel3')}
                      className="bg-purple-600 text-white px-4 py-2 rounded-lg hover:bg-purple-700 flex items-center gap-2"
                    >
                      <FileText size={16} /> Copy
                    </button>
                    <button
                      onClick={() =>
                        downloadTableAsExcel('tabel3', 'evaluasi_kehadiran')
                      }
                      className="bg-purple-600 text-white px-4 py-2 rounded-lg hover:bg-purple-700 flex items-center gap-2"
                    >
                      <Download size={16} /> Excel
                    </button>
                    <button
                      onClick={() =>
                        downloadAsPdf('tabel3', 'evaluasi_kehadiran')
                      }
                      className="bg-purple-600 text-white px-4 py-2 rounded-lg hover:bg-purple-700 flex items-center gap-2"
                    >
                      <Download size={16} /> PDF
                    </button>
                  </div>
                </div>
                <div className="overflow-x-auto border rounded-lg mb-4 sm:mb-6 md:mb-8 -mx-4 sm:mx-0">
                  <div className="min-w-max px-4 sm:px-0">
                    <table
                      id="tabel3"
                      className="min-w-full text-sm border-collapse bg-white"
                    >
                      <thead>
                        <tr className="bg-gray-300">
                          <th className="border border-gray-400 px-1.5 sm:px-2 py-1.5 sm:py-2 text-black font-bold text-xs sm:text-sm">
                            No
                          </th>
                          <th className="border border-gray-400 px-1.5 sm:px-2 py-1.5 sm:py-2 text-black font-bold text-xs sm:text-sm">
                            ID
                          </th>
                          <th
                            className="border border-gray-400 px-2 py-2 text-black font-bold"
                            style={{ minWidth: '220px' }}
                          >
                            Nama
                          </th>
                          <th className="border border-gray-400 px-6 py-3 text-black font-bold min-w-[120px]">
                            Jabatan
                          </th>
                          {recapData.dateRange.map((date) => (
                            <th
                              key={date}
                              className="border border-gray-400 px-2 py-2 text-black font-bold"
                            >
                              {getDateLabel(date)}
                            </th>
                          ))}
                          <th className="border border-gray-400 px-2 py-2 text-black font-bold bg-gray-300">
                            Hari Kerja
                          </th>
                          <th className="border border-gray-400 px-2 py-2 text-black font-bold bg-green-200">
                            Hadir
                          </th>
                          <th className="border border-gray-400 px-2 py-2 text-black font-bold bg-green-200">
                            Tepat
                          </th>
                          <th className="border border-gray-400 px-2 py-2 text-black font-bold bg-yellow-200">
                            Telat
                          </th>
                          <th className="border border-gray-400 px-2 py-2 text-black font-bold bg-red-200">
                            Alfa
                          </th>
                          <th className="border border-gray-400 px-2 py-2 text-black font-bold bg-purple-200">
                            %
                          </th>
                        </tr>
                      </thead>
                      <tbody>
                        {recapData.recap.map((emp) => {
                          let hariKerja = 0;
                          let totalH = 0;
                          let tepat = 0;
                          let telat = 0;
                          let alfa = 0;

                          recapData.dateRange.forEach((dateStr) => {
                            const ev = emp.dailyEvaluation[dateStr];
                            if (ev.text !== 'L') {
                              hariKerja++;
                              if (ev.text === 'H') {
                                totalH++;
                                if (ev.color === '90EE90') tepat++;
                                else if (ev.color === 'FFFF99') telat++;
                              } else if (ev.text === '-') {
                                alfa++;
                              }
                            }
                          });

                          const persentase =
                            hariKerja > 0
                              ? Math.round((totalH / hariKerja) * 100)
                              : 0;

                          return (
                            <tr key={emp.no}>
                              <td className="border border-gray-400 px-1.5 sm:px-2 py-1.5 sm:py-2 text-center text-xs sm:text-sm">
                                {emp.no}
                              </td>
                              <td className="border border-gray-400 px-1.5 sm:px-2 py-1.5 sm:py-2 text-center text-xs sm:text-sm">
                                {emp.id}
                              </td>
                              <td className="border border-gray-400 px-2 py-2">
                                {emp.name}
                              </td>
                              <td className="border border-gray-400 px-2 py-2">
                                {emp.position}
                              </td>
                              {recapData.dateRange.map((date) => {
                                const ev = emp.dailyEvaluation[date];
                                const bgColor = '#' + ev.color;
                                return (
                                  <td
                                    key={date}
                                    className="border border-gray-400 px-2 py-2 text-center font-bold"
                                    style={{ backgroundColor: bgColor }}
                                  >
                                    {ev.text}
                                  </td>
                                );
                              })}
                              <td className="border border-gray-400 px-2 py-2 text-center font-bold bg-gray-100">
                                {hariKerja}
                              </td>
                              <td className="border border-gray-400 px-2 py-2 text-center font-bold bg-green-100">
                                {totalH}
                              </td>
                              <td className="border border-gray-400 px-2 py-2 text-center font-bold bg-green-100">
                                {tepat}
                              </td>
                              <td className="border border-gray-400 px-2 py-2 text-center font-bold bg-yellow-100">
                                {telat}
                              </td>
                              <td className="border border-gray-400 px-2 py-2 text-center font-bold bg-red-100">
                                {alfa}
                              </td>
                              <td className="border border-gray-400 px-2 py-2 text-center font-bold bg-purple-100">
                                {persentase}%
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
              <div
                ref={panduanRef}
                className="mt-4 sm:mt-6 p-4 sm:p-6 bg-gradient-to-br from-blue-50 to-indigo-50 border-2 border-blue-200 rounded-lg sm:rounded-xl"
              >
                <div className="max-w-6xl mx-auto p-3 sm:p-4 md:p-6 bg-gray-50">
                  <h4 className="text-2xl font-bold text-gray-800 mb-6 flex items-center gap-2">
                    📊 Panduan Lengkap 3 Tabel Rekap
                  </h4>

                  {/* TABEL 1: Rekap Mesin */}
                  <div className="mb-6 bg-white p-6 rounded-lg shadow-sm border-l-4 border-blue-500">
                    <h5 className="font-bold text-xl text-blue-700 mb-3">
                      1️⃣ Rekap Mesin (Data Mentah)
                    </h5>
                    <p className="text-gray-700 mb-5">
                      Data langsung dari mesin fingerprint tanpa proses evaluasi
                    </p>

                    <div className="grid md:grid-cols-2 gap-4">
                      {/* Biru */}
                      <div className="bg-blue-50 p-4 rounded-lg border border-blue-200">
                        <div className="flex items-start gap-3">
                          <div className="flex-shrink-0">
                            <div className="w-20 h-16 bg-blue-300 rounded-lg flex flex-col items-center justify-center text-xs font-bold">
                              <div>07:15</div>
                              <div className="text-blue-800">16:30</div>
                            </div>
                          </div>
                          <div className="flex-1">
                            <span className="font-semibold text-blue-800 text-base">
                              🔵 BIRU = Scan Lengkap
                            </span>
                            <p className="text-sm text-gray-700 mt-1">
                              Scan <strong>MASUK dan PULANG</strong> keduanya
                              tercatat
                            </p>
                            <p className="text-xs text-blue-700 mt-1 italic">
                              ✅ Data absensi lengkap
                            </p>
                          </div>
                        </div>
                      </div>

                      {/* Kuning */}
                      <div className="bg-yellow-50 p-4 rounded-lg border border-yellow-200">
                        <div className="flex items-start gap-3">
                          <div className="flex-shrink-0">
                            <div className="w-20 h-16 bg-yellow-300 rounded-lg flex flex-col items-center justify-center text-xs font-bold">
                              <div>07:15</div>
                              <div className="text-yellow-800">-</div>
                            </div>
                          </div>
                          <div className="flex-1">
                            <span className="font-semibold text-yellow-800 text-base">
                              🟡 KUNING = Scan Tidak Lengkap
                            </span>
                            <p className="text-sm text-gray-700 mt-1">
                              Hanya scan <strong>MASUK saja</strong> atau{' '}
                              <strong>PULANG saja</strong>
                            </p>
                            <p className="text-xs text-yellow-700 mt-1 italic">
                              ⚠️ Data absensi tidak lengkap
                            </p>
                          </div>
                        </div>
                      </div>

                      {/* Merah */}
                      <div className="bg-red-50 p-4 rounded-lg border border-red-200">
                        <div className="flex items-start gap-3">
                          <div className="flex-shrink-0">
                            <div className="w-20 h-16 bg-red-300 rounded-lg flex flex-col items-center justify-center text-xs font-bold">
                              <div>-</div>
                              <div className="text-red-800">-</div>
                            </div>
                          </div>
                          <div className="flex-1">
                            <span className="font-semibold text-red-800 text-base">
                              🔴 MERAH = Tidak Scan
                            </span>
                            <p className="text-sm text-gray-700 mt-1">
                              Tidak ada scan{' '}
                              <strong>MASUK maupun PULANG</strong> (Alpha)
                            </p>
                            <p className="text-xs text-red-700 mt-1 italic">
                              ❌ Tidak ada data absensi
                            </p>
                          </div>
                        </div>
                      </div>

                      {/* Putih L */}
                      <div className="bg-gray-50 p-4 rounded-lg border border-gray-300">
                        <div className="flex items-start gap-3">
                          <div className="flex-shrink-0">
                            <div className="w-20 h-16 bg-white border-2 border-gray-400 rounded-lg flex items-center justify-center text-2xl font-bold text-gray-700">
                              L
                            </div>
                          </div>
                          <div className="flex-1">
                            <span className="font-semibold text-gray-800 text-base">
                              ⚪ PUTIH + L = Libur
                            </span>
                            <p className="text-sm text-gray-700 mt-1">
                              Hari <strong>LIBUR</strong> atau tidak ada jadwal
                              (Jumat/OFF)
                            </p>
                            <p className="text-xs text-gray-600 mt-1 italic">
                              📅 Tidak perlu scan
                            </p>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>

                  {/* TABEL 2: Kedisiplinan Waktu */}
                  <div className="mb-6 bg-white p-6 rounded-lg shadow-sm border-l-4 border-green-500">
                    <h5 className="font-bold text-xl text-green-700 mb-3">
                      2️⃣ Kedisiplinan Waktu (Evaluasi Disiplin)
                    </h5>
                    <p className="text-gray-700 mb-5">
                      Evaluasi kedisiplinan berdasarkan jadwal mengajar dan
                      kelengkapan scan
                    </p>

                    <div className="grid md:grid-cols-2 gap-4">
                      {/* Hijau */}
                      <div className="bg-green-50 p-4 rounded-lg border border-green-200">
                        <div className="flex items-start gap-3">
                          <div className="flex-shrink-0">
                            <div className="w-20 h-16 bg-green-300 rounded-lg flex flex-col items-center justify-center text-xs font-bold">
                              <div>Jadwal: 07:00</div>
                              <div className="text-green-800">
                                Scan: 06:45 ✓
                              </div>
                            </div>
                          </div>
                          <div className="flex-1">
                            <span className="font-semibold text-green-800 text-base">
                              🟢 HIJAU = Disiplin Tinggi
                            </span>
                            <p className="text-sm text-gray-700 mt-1">
                              Datang <strong>SEBELUM</strong> jadwal mengajar
                              dimulai
                            </p>
                            <p className="text-xs text-green-700 mt-1 italic">
                              ✨ Guru paling disiplin!
                            </p>
                          </div>
                        </div>
                      </div>

                      {/* Biru */}
                      <div className="bg-blue-50 p-4 rounded-lg border border-blue-200">
                        <div className="flex items-start gap-3">
                          <div className="flex-shrink-0">
                            <div className="w-20 h-16 bg-blue-300 rounded-lg flex flex-col items-center justify-center text-xs font-bold">
                              <div>Masuk: 07:15</div>
                              <div className="text-blue-800">
                                Pulang: 16:30 ✓
                              </div>
                            </div>
                          </div>
                          <div className="flex-1">
                            <span className="font-semibold text-blue-800 text-base">
                              🔵 BIRU = Scan Lengkap
                            </span>
                            <p className="text-sm text-gray-700 mt-1">
                              Scan <strong>MASUK dan PULANG</strong> lengkap
                              (administrasi tertib)
                            </p>
                            <p className="text-xs text-blue-700 mt-1 italic">
                              📋 Tertib administrasi
                            </p>
                          </div>
                        </div>
                      </div>

                      {/* Kuning */}
                      <div className="bg-yellow-50 p-4 rounded-lg border border-yellow-200">
                        <div className="flex items-start gap-3">
                          <div className="flex-shrink-0">
                            <div className="w-20 h-16 bg-yellow-300 rounded-lg flex flex-col items-center justify-center text-xs font-bold">
                              <div>Masuk: 08:15</div>
                              <div className="text-yellow-800">Pulang: - ✗</div>
                            </div>
                          </div>
                          <div className="flex-1">
                            <span className="font-semibold text-yellow-800 text-base">
                              🟡 KUNING = Tidak Lengkap
                            </span>
                            <p className="text-sm text-gray-700 mt-1">
                              Datang tapi <strong>LUPA scan pulang</strong>,
                              atau datang setelah jadwal
                            </p>
                            <p className="text-xs text-yellow-700 mt-1 italic">
                              ⚠️ Perlu lebih teliti scan
                            </p>
                          </div>
                        </div>
                      </div>

                      {/* Merah */}
                      <div className="bg-red-50 p-4 rounded-lg border border-red-200">
                        <div className="flex items-start gap-3">
                          <div className="flex-shrink-0">
                            <div className="w-20 h-16 bg-red-300 rounded-lg flex flex-col items-center justify-center text-xs font-bold">
                              <div>Masuk: -</div>
                              <div className="text-red-800">Pulang: -</div>
                            </div>
                          </div>
                          <div className="flex-1">
                            <span className="font-semibold text-red-800 text-base">
                              🔴 MERAH = Tidak Absen / Alfa
                            </span>
                            <p className="text-sm text-gray-700 mt-1">
                              Datang tapi tidak scan sama sekali, atau tidak
                              hadir
                            </p>
                            <p className="text-xs text-red-700 mt-1 italic">
                              ❌ Kesadaran absensi kurang
                            </p>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>

                  {/* TABEL 3: Evaluasi Kehadiran */}
                  <div className="mb-6 bg-white p-6 rounded-lg shadow-sm border-l-4 border-purple-500">
                    <h5 className="font-bold text-xl text-purple-700 mb-3">
                      3️⃣ Evaluasi Kehadiran (Status Resmi)
                    </h5>
                    <p className="text-gray-700 mb-5">
                      Status kehadiran berdasarkan jadwal mengajar dan ketepatan
                      waktu
                    </p>

                    <div className="grid md:grid-cols-2 gap-4">
                      {/* H Hijau */}
                      <div className="bg-green-50 p-4 rounded-lg border border-green-200">
                        <div className="flex items-start gap-3">
                          <div className="flex-shrink-0">
                            <div className="w-20 h-16 bg-green-300 rounded-lg flex items-center justify-center text-3xl font-bold text-green-800">
                              H
                            </div>
                          </div>
                          <div className="flex-1">
                            <span className="font-semibold text-green-800 text-base">
                              🟢 H HIJAU = Hadir Tepat Waktu
                            </span>
                            <p className="text-sm text-gray-700 mt-1">
                              Scan <strong>SEBELUM/TEPAT</strong> jadwal
                              mengajar dimulai
                            </p>
                            <p className="text-xs text-green-700 mt-1 italic">
                              ✅ Status kehadiran sempurna
                            </p>
                          </div>
                        </div>
                      </div>

                      {/* H Kuning */}
                      <div className="bg-yellow-50 p-4 rounded-lg border border-yellow-200">
                        <div className="flex items-start gap-3">
                          <div className="flex-shrink-0">
                            <div className="w-20 h-16 bg-yellow-300 rounded-lg flex items-center justify-center text-3xl font-bold text-yellow-800">
                              H
                            </div>
                          </div>
                          <div className="flex-1">
                            <span className="font-semibold text-yellow-800 text-base">
                              🟡 H KUNING = Hadir Terlambat
                            </span>
                            <p className="text-sm text-gray-700 mt-1">
                              Scan <strong>SETELAH</strong> jadwal mengajar
                              dimulai
                            </p>
                            <p className="text-xs text-yellow-700 mt-1 italic">
                              ⚠️ Status kehadiran terlambat
                            </p>
                          </div>
                        </div>
                      </div>

                      {/* Strip Merah */}
                      <div className="bg-red-50 p-4 rounded-lg border border-red-200">
                        <div className="flex items-start gap-3">
                          <div className="flex-shrink-0">
                            <div className="w-20 h-16 bg-red-300 rounded-lg flex items-center justify-center text-4xl font-bold text-red-800">
                              -
                            </div>
                          </div>
                          <div className="flex-1">
                            <span className="font-semibold text-red-800 text-base">
                              🔴 STRIP = Alpha
                            </span>
                            <p className="text-sm text-gray-700 mt-1">
                              Tidak hadir padahal <strong>ADA JADWAL</strong>{' '}
                              mengajar
                            </p>
                            <p className="text-xs text-red-700 mt-1 italic">
                              ❌ Status tidak hadir
                            </p>
                          </div>
                        </div>
                      </div>

                      {/* L Putih */}
                      <div className="bg-gray-50 p-4 rounded-lg border border-gray-300">
                        <div className="flex items-start gap-3">
                          <div className="flex-shrink-0">
                            <div className="w-20 h-16 bg-white border-2 border-gray-400 rounded-lg flex items-center justify-center text-3xl font-bold text-gray-700">
                              L
                            </div>
                          </div>
                          <div className="flex-1">
                            <span className="font-semibold text-gray-800 text-base">
                              ⚪ L = Libur
                            </span>
                            <p className="text-sm text-gray-700 mt-1">
                              Tidak ada <strong>JADWAL MENGAJAR</strong>{' '}
                              (OFF/Jumat)
                            </p>
                            <p className="text-xs text-gray-600 mt-1 italic">
                              📅 Status hari libur
                            </p>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>

                  {/* Tips Membaca */}
                  <div className="mb-6 p-5 bg-gradient-to-r from-indigo-100 to-purple-100 rounded-lg border border-indigo-300">
                    <h6 className="font-bold text-lg text-indigo-800 mb-3">
                      💡 Tips Membaca Rekap:
                    </h6>
                    <ul className="text-sm text-gray-700 space-y-2">
                      <li className="flex items-start gap-2">
                        <span className="font-bold text-indigo-600 min-w-[70px]">
                          Tabel 1
                        </span>
                        <span>
                          untuk cek kehadiran mentah (ada scan atau tidak)
                        </span>
                      </li>
                      <li className="flex items-start gap-2">
                        <span className="font-bold text-indigo-600 min-w-[70px]">
                          Tabel 2
                        </span>
                        <span>
                          untuk evaluasi kedisiplinan (tepat waktu atau lupa
                          scan pulang)
                        </span>
                      </li>
                      <li className="flex items-start gap-2">
                        <span className="font-bold text-indigo-600 min-w-[70px]">
                          Tabel 3
                        </span>
                        <span>untuk status resmi kehadiran (H/L/-)</span>
                      </li>
                      <li className="flex items-start gap-2">
                        <span className="font-bold text-indigo-600 min-w-[70px]">
                          Kolom Kanan
                        </span>
                        <span>
                          di setiap tabel menunjukkan{' '}
                          <strong>total perhitungan</strong>
                        </span>
                      </li>
                    </ul>
                  </div>

                  {/* Tombol Action */}
                  <div className="flex justify-end gap-2 sm:gap-4">
                    <button
                      onClick={handleCopyPanduan}
                      className="px-3 py-2 sm:px-5 sm:py-3 
               bg-green-500 text-white rounded-lg 
               hover:bg-green-600 font-medium 
               text-xs sm:text-sm 
               flex items-center gap-1.5 sm:gap-2 
               shadow-md hover:shadow-lg transition-all"
                    >
                      <FileText size={18} className="sm:size-20" />
                      Copy JPG
                    </button>

                    <button
                      onClick={handleDownloadPanduan}
                      className="px-3 py-2 sm:px-5 sm:py-3 
               bg-blue-500 text-white rounded-lg 
               hover:bg-blue-600 font-medium 
               text-xs sm:text-sm 
               flex items-center gap-1.5 sm:gap-2 
               shadow-md hover:shadow-lg transition-all"
                    >
                      <Download size={18} className="sm:size-20" />
                      Download JPG
                    </button>
                  </div>
                </div>
              </div>
            </div>
          )}
          <div className="mt-4 sm:mt-6 md:mt-8 bg-blue-50 rounded-lg sm:rounded-xl p-4 sm:p-6">
            <h3 className="font-semibold text-base sm:text-lg text-gray-800 mb-3">
              Cara Penggunaan:
            </h3>
            <ol className="list-decimal list-inside space-y-2 text-gray-700 text-xs sm:text-sm">
              <li>Upload Laporan Absensi</li>
              <li>Upload Jadwal Kerja (opsional)</li>
              <li>Pilih tanggal awal dan akhir</li>
              <li>Klik Generate Tabel Rekap</li>
              <li>Klik Generate Kesimpulan Profil</li>
              <li>
                Klik Copy atau Download Excel atau Download PDF pada tabel yang
                diinginkan
              </li>
            </ol>
          </div>
        </div>
      </div>
    </div>
  );
};

export default AttendanceRecapSystem;
