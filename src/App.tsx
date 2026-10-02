import React, { useState, useEffect, useRef } from "react";
import jsQR from "jsqr";
// @ts-ignore
import QRCode from "qrcode";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
// @ts-ignore
import { Canvg } from "canvg";
import SignatureCanvas from "react-signature-canvas";
import {
  Chart as ChartJS,
  CategoryScale,
  LinearScale,
  BarElement,
  Title,
  Tooltip,
  Legend,
  ChartData,
  ChartOptions,
  ChartEvent,
  LegendItem,
} from "chart.js";
import { Bar } from "react-chartjs-2";
import {
  configureSync,
  syncAll,
  getStudentsLocal,
  getAttendanceByDate,
  getFotoAbsenLocal,
  getJadwalLocal,
  getTanggalMerahLocal,
} from "./offlineDb";

ChartJS.register(
  CategoryScale,
  LinearScale,
  BarElement,
  Title,
  Tooltip,
  Legend
);

const endpoint =
  "https://script.google.com/macros/s/AKfycby9kyvmzPfwGE-C0cxEaxKQx7wSLvB7djEhx48JoRxjW6InKmg3qg4aad-3XU7OgCnp/exec";
configureSync(endpoint);
const SHEET_SEMESTER1 = "RekapSemester1";
const SHEET_SEMESTER2 = "RekapSemester2";

interface Student {
  id: string;
  name: string;
  nisn: string; // ← Ubah dari string | null | undefined menjadi string
  kelas: string;
  jenisKelamin: string;
  foto?: string;
}

interface SchoolData {
  namaKepsek: string;
  nipKepsek: string;
  ttdKepsek: string;
  namaGuru: string;
  nipGuru: string;
  ttdGuru: string;
  namaKota: string;
  statusGuru: string;
  namaSekolah: string;
}

type AttendanceStatus = "Hadir" | "Izin" | "Sakit" | "Alpha";

interface AttendanceRecord {
  [date: string]: {
    [studentId: string]: AttendanceStatus;
  };
}

interface KeteranganRecord {
  [date: string]: {
    [studentId: string]: string;
  };
}

interface MonthlyRecap {
  nama: string;
  kelas: string;
  hadir: number;
  alpa: number;
  izin: number;
  sakit: number;
  persenHadir: number;
}

interface GraphData {
  [month: string]: {
    Hadir: number;
    Alpha: number;
    Izin: number;
    Sakit: number;
  };
}

interface StatusSummary {
  Hadir: number;
  Izin: number;
  Sakit: number;
  Alpha: number;
}

interface StatusVisibility {
  Hadir: boolean;
  Alpha: boolean;
  Izin: boolean;
  Sakit: boolean;
}

interface AttendanceHistory {
  tanggal: string;
  nama: string;
  kelas: string;
  nisn: string;
  status: AttendanceStatus;
}

interface SemesterRecap {
  nama: string;
  kelas: string;
  hadir: number;
  alpa: number;
  izin: number;
  sakit: number;
  persenHadir: number;
}

interface TanggalMerah {
  tanggal: string;
  deskripsi: string;
  tanggalAkhir?: string;
}

interface JadwalMengajar {
  kelas: string;
  hari: string;
}

const formatDateDDMMYYYY = (isoDate: string): string => {
  const [year, month, day] = isoDate.split("-");
  return `${day}/${month}/${year}`;
};

type EditedRecord = {
  date: string;
  nisn: string;
  status: AttendanceStatus | "";
};

const SchoolDataTab: React.FC<{
  onRefresh: () => void;
}> = ({ onRefresh }) => {
  const [schoolData, setSchoolData] = useState<SchoolData | null>(null);
  const [namaKepsek, setNamaKepsek] = useState("");
  const [nipKepsek, setNipKepsek] = useState("");
  const [namaGuru, setNamaGuru] = useState("");
  const [nipGuru, setNipGuru] = useState("");
  const [ttdKepsek, setTtdKepsek] = useState("");
  const [ttdGuru, setTtdGuru] = useState("");
  const [loading, setLoading] = useState<boolean>(true);
  const [isSaving, setIsSaving] = useState<boolean>(false); // Add isSaving state
  const [isKepsekSigning, setIsKepsekSigning] = useState(false);
  const [isGuruSigning, setIsGuruSigning] = useState(false);
  const kepsekSigCanvas = useRef<SignatureCanvas>(null);
  const guruSigCanvas = useRef<SignatureCanvas>(null);
  const [namaKota, setNamaKota] = useState("");
  const [statusGuru, setStatusGuru] = useState("Guru Kelas");
  const [namaSekolah, setNamaSekolah] = useState("");

  useEffect(() => {
    fetch(`${endpoint}?action=schoolData`)
      .then((res) => {
        if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
        return res.json();
      })
      .then((data) => {
        if (data.success && data.data && data.data.length > 0) {
          const record = data.data[0];
          setSchoolData(record);
          setNamaKepsek(record.namaKepsek);
          setNipKepsek(record.nipKepsek);
          setTtdKepsek(record.ttdKepsek);
          setNamaGuru(record.namaGuru);
          setNipGuru(record.nipGuru);
          setTtdGuru(record.ttdGuru);
          setNamaKota(record.namaKota);
          setStatusGuru(record.statusGuru || "Guru Kelas");
          setNamaSekolah(record.namaSekolah || "");
        } else {
          setSchoolData(null);
        }
        setLoading(false);
      })
      .catch((error) => {
        console.error("Error fetching school data:", error);
        alert("❌ Gagal memuat data sekolah. Cek console untuk detail.");
        setLoading(false);
      });
  }, []);

  const handleSave = () => {
    if (!namaKepsek || !nipKepsek || !namaGuru || !nipGuru) {
      alert("⚠️ Nama dan NIP Kepala Sekolah serta Guru wajib diisi!");
      return;
    }

    setIsSaving(true);

    const data: SchoolData = {
      namaKepsek,
      nipKepsek,
      ttdKepsek: ttdKepsek || "", // Sudah benar - bisa kosong
      namaGuru,
      nipGuru,
      ttdGuru: ttdGuru || "", // Sudah benar - bisa kosong
      namaKota: namaKota || "",
      statusGuru: statusGuru || "Guru Kelas",
      namaSekolah: namaSekolah || "",
    };

    fetch(endpoint, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        type: "schoolData",
        ...data,
      }),
    })
      .then(() => {
        alert("✅ Data sekolah berhasil diperbarui!");
        onRefresh();
        setIsSaving(false); // Reset saving state on success
      })
      .catch(() => {
        alert("❌ Gagal memperbarui data sekolah.");
        setIsSaving(false); // Reset saving state on error
      });
  };

  const handleClearKepsekSignature = () => {
    kepsekSigCanvas.current?.clear();
  };

  const handleClearGuruSignature = () => {
    guruSigCanvas.current?.clear();
  };

  const handleSaveKepsekSignature = () => {
    const signature = kepsekSigCanvas.current?.toDataURL("image/svg+xml"); // Ubah ke SVG
    if (signature && !kepsekSigCanvas.current?.isEmpty()) {
      setTtdKepsek(signature);
      setIsKepsekSigning(false);
    } else {
      alert("⚠️ Tanda tangan kepala sekolah kosong!");
    }
  };

  const handleSaveGuruSignature = () => {
    const signature = guruSigCanvas.current?.toDataURL("image/svg+xml"); // Ubah ke SVG
    if (signature && !guruSigCanvas.current?.isEmpty()) {
      setTtdGuru(signature);
      setIsGuruSigning(false);
    } else {
      alert("⚠️ Tanda tangan guru kosong!");
    }
  };

  const handleStartKepsekSigning = () => {
    setIsKepsekSigning(true);
    kepsekSigCanvas.current?.clear();
  };

  const handleStartGuruSigning = () => {
    setIsGuruSigning(true);
    guruSigCanvas.current?.clear();
  };

  if (loading) {
    return (
      <div className="text-center py-8">
        <p className="text-gray-500">Memuat data sekolah...</p>
      </div>
    );
  }

  return (
    <div className="max-w-4xl mx-auto" style={{ paddingBottom: "70px" }}>
      <div className="bg-white p-6 rounded-lg shadow-md">
        <h2 className="text-2xl font-bold text-center text-blue-700 mb-6">
          🏫 Data Sekolah
        </h2>
        <div className="grid grid-cols-1 gap-4 mb-6">
          <div>
            <h3 className="text-lg font-semibold text-gray-700 mb-2">
              Nama Sekolah
            </h3>
            <input
              type="text"
              placeholder="Nama Sekolah"
              value={namaSekolah}
              onChange={(e) => setNamaSekolah(e.target.value)}
              className="w-full border border-gray-300 px-4 py-2 rounded-lg mb-2"
              disabled={isSaving}
            />
          </div>
          <div>
            <h3 className="text-lg font-semibold text-gray-700 mb-2">
              Kota/Kabupaten
            </h3>
            <input
              type="text"
              placeholder="Nama Kota/Kabupaten"
              value={namaKota}
              onChange={(e) => setNamaKota(e.target.value)}
              className="w-full border border-gray-300 px-4 py-2 rounded-lg mb-2"
              disabled={isSaving}
            />
          </div>
          <div>
            <h3 className="text-lg font-semibold text-gray-700 mb-2">
              Status Guru
            </h3>
            <select
              value={statusGuru}
              onChange={(e) => setStatusGuru(e.target.value)}
              className="w-full border border-gray-300 px-4 py-2 rounded-lg mb-2"
              disabled={isSaving}
            >
              <option value="Guru Kelas">Guru Kelas</option>
              <option value="Guru PJOK">Guru PJOK</option>
              <option value="Guru PAI">Guru PAI</option>
            </select>
          </div>
          <div>
            <h3 className="text-lg font-semibold text-gray-700 mb-2">
              Kepala Sekolah
            </h3>
            <input
              type="text"
              placeholder="Nama Kepala Sekolah"
              value={namaKepsek}
              onChange={(e) => setNamaKepsek(e.target.value)}
              className="w-full border border-gray-300 px-4 py-2 rounded-lg mb-2"
              disabled={isSaving} // Disable input during saving
            />
            <input
              type="text"
              placeholder="NIP Kepala Sekolah"
              value={nipKepsek}
              onChange={(e) => setNipKepsek(e.target.value)}
              className="w-full border border-gray-300 px-4 py-2 rounded-lg mb-2"
              disabled={isSaving} // Disable input during saving
            />
            <div className="mb-2">
              <p className="text-sm text-gray-500 mb-1">
                Tanda Tangan Kepala Sekolah
              </p>
              <div className="relative">
                <SignatureCanvas
                  ref={kepsekSigCanvas}
                  penColor="black"
                  canvasProps={{
                    className: `border border-gray-300 rounded-lg ${
                      !isKepsekSigning || isSaving
                        ? "opacity-50 pointer-events-none"
                        : ""
                    }`,
                    style: { width: "100%", height: "300px" },
                  }}
                  clearOnResize={false}
                />
                {!isKepsekSigning && (
                  <div className="absolute inset-0 flex items-center justify-center bg-gray-200 bg-opacity-50">
                    <span className="text-gray-500">
                      Klik "Mulai Tanda Tangan" untuk mengaktifkan
                    </span>
                  </div>
                )}
              </div>
              <div className="flex gap-2 mt-2">
                {!isKepsekSigning && (
                  <button
                    onClick={handleStartKepsekSigning}
                    className={`px-4 py-1 bg-green-500 hover:bg-green-600 text-white rounded-lg text-sm ${
                      isSaving ? "opacity-50 cursor-not-allowed" : ""
                    }`}
                    disabled={isSaving} // Disable button during saving
                  >
                    ✍️ Mulai Tanda Tangan
                  </button>
                )}
                {isKepsekSigning && (
                  <button
                    onClick={handleSaveKepsekSignature}
                    className={`px-4 py-1 bg-blue-500 hover:bg-blue-600 text-white rounded-lg text-sm ${
                      isSaving ? "opacity-50 cursor-not-allowed" : ""
                    }`}
                    disabled={isSaving} // Disable button during saving
                  >
                    💾 Simpan Tanda Tangan
                  </button>
                )}
                <button
                  onClick={handleClearKepsekSignature}
                  className={`px-4 py-1 bg-red-500 hover:bg-red-600 text-white rounded-lg text-sm ${
                    !isKepsekSigning || isSaving
                      ? "opacity-50 cursor-not-allowed"
                      : ""
                  }`}
                  disabled={!isKepsekSigning || isSaving} // Disable during saving
                >
                  🗑️ Hapus TTD
                </button>
              </div>
            </div>
            {ttdKepsek && (
              <>
                <img
                  src={ttdKepsek}
                  alt="Tanda Tangan Kepala Sekolah"
                  className="mt-2 max-w-full h-20 border border-gray-200 rounded-lg"
                />
                <button
                  onClick={() => {
                    if (
                      confirm(
                        "Hapus tanda tangan Kepala Sekolah yang tersimpan?"
                      )
                    ) {
                      setTtdKepsek("");
                    }
                  }}
                  disabled={isSaving}
                  className={`mt-2 px-4 py-1 bg-red-500 hover:bg-red-600 text-white rounded-lg text-sm ${
                    isSaving ? "opacity-50 cursor-not-allowed" : ""
                  }`}
                >
                  🗑️ Hapus TTD Tersimpan
                </button>
              </>
            )}
          </div>
          <div>
            <h3 className="text-lg font-semibold text-gray-700 mb-2">Guru</h3>
            <input
              type="text"
              placeholder="Nama Guru"
              value={namaGuru}
              onChange={(e) => setNamaGuru(e.target.value)}
              className="w-full border border-gray-300 px-4 py-2 rounded-lg mb-2 text-gray-400 bg-gray-50"
              disabled={true}
            />
            <input
              type="text"
              placeholder="NIP Guru"
              value={nipGuru}
              onChange={(e) => setNipGuru(e.target.value)}
              className="w-full border border-gray-300 px-4 py-2 rounded-lg mb-2 text-gray-400 bg-gray-50"
              disabled={true}
            />
            <div className="mb-2">
              <p className="text-sm text-gray-500 mb-1">Tanda Tangan Guru</p>
              <div className="relative">
                <SignatureCanvas
                  ref={guruSigCanvas}
                  penColor="black"
                  canvasProps={{
                    className: `border border-gray-300 rounded-lg ${
                      !isGuruSigning || isSaving
                        ? "opacity-50 pointer-events-none"
                        : ""
                    }`,
                    style: { width: "100%", height: "300px" },
                  }}
                  clearOnResize={false}
                />
                {!isGuruSigning && (
                  <div className="absolute inset-0 flex items-center justify-center bg-gray-200 bg-opacity-50">
                    <span className="text-gray-500">
                      Klik "Mulai Tanda Tangan" untuk mengaktifkan
                    </span>
                  </div>
                )}
              </div>
              <div className="flex gap-2 mt-2">
                {!isGuruSigning && (
                  <button
                    onClick={handleStartGuruSigning}
                    className={`px-4 py-1 bg-green-500 hover:bg-green-600 text-white rounded-lg text-sm ${
                      isSaving ? "opacity-50 cursor-not-allowed" : ""
                    }`}
                    disabled={isSaving} // Disable button during saving
                  >
                    ✍️ Mulai Tanda Tangan
                  </button>
                )}
                {isGuruSigning && (
                  <button
                    onClick={handleSaveGuruSignature}
                    className={`px-4 py-1 bg-blue-500 hover:bg-blue-600 text-white rounded-lg text-sm ${
                      isSaving ? "opacity-50 cursor-not-allowed" : ""
                    }`}
                    disabled={isSaving} // Disable button during saving
                  >
                    💾 Simpan Tanda Tangan
                  </button>
                )}
                <button
                  onClick={handleClearGuruSignature}
                  className={`px-4 py-1 bg-red-500 hover:bg-red-600 text-white rounded-lg text-sm ${
                    !isGuruSigning || isSaving
                      ? "opacity-50 cursor-not-allowed"
                      : ""
                  }`}
                  disabled={!isGuruSigning || isSaving} // Disable during saving
                >
                  🗑️ Hapus TTD
                </button>
              </div>
            </div>
            {ttdGuru && (
              <>
                <img
                  src={ttdGuru}
                  alt="Tanda Tangan Guru"
                  className="mt-2 max-w-full h-20 border border-gray-200 rounded-lg"
                />
                <button
                  onClick={() => {
                    if (confirm("Hapus tanda tangan Guru yang tersimpan?")) {
                      setTtdGuru("");
                    }
                  }}
                  disabled={isSaving}
                  className={`mt-2 px-4 py-1 bg-red-500 hover:bg-red-600 text-white rounded-lg text-sm ${
                    isSaving ? "opacity-50 cursor-not-allowed" : ""
                  }`}
                >
                  🗑️ Hapus TTD Tersimpan
                </button>
              </>
            )}
          </div>
        </div>
        <div className="text-center">
          <button
            onClick={handleSave}
            disabled={isSaving} // Disable button during saving
            className={`px-6 py-2 rounded-lg font-medium transition-colors ${
              isSaving
                ? "bg-blue-400 cursor-not-allowed"
                : "bg-blue-600 hover:bg-blue-700"
            } text-white`}
          >
            {isSaving ? "⏳ Menyimpan..." : "💾 Simpan Data Sekolah"}
          </button>
        </div>
      </div>
    </div>
  );
};

const StudentDataTab: React.FC<{
  students: Student[];
  onRefresh: () => void;
  uniqueClasses: string[];
  loading: boolean;
}> = ({ students, onRefresh, uniqueClasses, loading }) => {
  const [nisn, setNisn] = useState("");
  const [nama, setNama] = useState("");
  const [kelas, setKelas] = useState("");
  const [jenisKelamin, setJenisKelamin] = useState("");
  const [searchQuery, setSearchQuery] = useState("");
  const [selectedKelas, setSelectedKelas] = useState<string>("Semua");

  // State untuk bulk import
  const [showBulkImport, setShowBulkImport] = useState(false);
  const [bulkNisn, setBulkNisn] = useState("");
  const [bulkNama, setBulkNama] = useState("");
  const [bulkKelas, setBulkKelas] = useState("");
  const [bulkJenisKelamin, setBulkJenisKelamin] = useState("");

  // State untuk loading
  const [isSaving, setIsSaving] = useState(false);
  const [isBulkSaving, setIsBulkSaving] = useState(false);
  const [isEditing, setIsEditing] = useState(false);
  const [isDeleting, setIsDeleting] = useState(false);
  const [isImporting, setIsImporting] = useState(false);
  const [uploadingFotoNisn, setUploadingFotoNisn] = useState<string | null>(
    null
  );
  const [showFotoModal, setShowFotoModal] = useState(false);
  const [fotoModalStudent, setFotoModalStudent] = useState<Student | null>(
    null
  );
  const fotoInputRef = useRef<HTMLInputElement>(null);
  const [selectedFotoStudentNisn, setSelectedFotoStudentNisn] = useState<
    string | null
  >(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // QR Code States
  const [showQRModal, setShowQRModal] = useState(false);
  const [selectedStudent, setSelectedStudent] = useState<Student | null>(null);
  const [qrCodeDataURL, setQrCodeDataURL] = useState("");
  const [isGeneratingQR, setIsGeneratingQR] = useState(false);
  const [showBulkQRModal, setShowBulkQRModal] = useState(false);
  const [bulkQRCodes, setBulkQRCodes] = useState<
    Array<{ student: Student; qrCode: string }>
  >([]);
  const [isGeneratingBulkQR, setIsGeneratingBulkQR] = useState(false);

  // Tambahkan helper function ini di dalam StudentDataTab
  const getFotoUrl = (fotoUrl: string | undefined): string => {
    if (!fotoUrl) return "";

    // Jika sudah format lh3, langsung return
    if (fotoUrl.includes("lh3.googleusercontent.com")) return fotoUrl;

    // Konversi dari format uc?export=view ke lh3
    if (fotoUrl.includes("uc?export=view&id=")) {
      const fileId = new URLSearchParams(fotoUrl.split("?")[1]).get("id");
      if (fileId) return `https://lh3.googleusercontent.com/d/${fileId}`;
    }

    // Konversi dari format /file/d/ID/view
    if (fotoUrl.includes("/file/d/")) {
      const fileId = fotoUrl.split("/file/d/")[1]?.split("/")[0];
      if (fileId) return `https://lh3.googleusercontent.com/d/${fileId}`;
    }

    return fotoUrl;
  };

  const handleSubmit = () => {
    if (!nisn || !nama || !kelas || !jenisKelamin) {
      alert("⚠️ Semua field wajib diisi!");
      return;
    }

    setIsSaving(true);

    fetch(endpoint, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        type: "siswa",
        nisn,
        nama,
        kelas,
        jenisKelamin,
      }),
    })
      .then(() => {
        alert("✅ Siswa berhasil ditambahkan!");
        setNisn("");
        setNama("");
        setKelas("");
        setJenisKelamin("");
        onRefresh();
        setIsSaving(false);
      })
      .catch(() => {
        alert("❌ Gagal menambahkan siswa.");
        setIsSaving(false);
      });
  };

  const handleBulkImport = () => {
    if (
      !bulkNisn.trim() ||
      !bulkNama.trim() ||
      !bulkKelas.trim() ||
      !bulkJenisKelamin.trim()
    ) {
      alert("⚠️ Semua field data massal wajib diisi!");
      return;
    }

    const nisnLines = bulkNisn
      .trim()
      .split("\n")
      .filter((line) => line.trim());
    const namaLines = bulkNama
      .trim()
      .split("\n")
      .filter((line) => line.trim());
    const kelasLines = bulkKelas
      .trim()
      .split("\n")
      .filter((line) => line.trim());
    const jenisKelaminLines = bulkJenisKelamin
      .trim()
      .split("\n")
      .filter((line) => line.trim());

    if (
      nisnLines.length !== namaLines.length ||
      namaLines.length !== kelasLines.length ||
      kelasLines.length !== jenisKelaminLines.length
    ) {
      alert(
        "⚠️ Jumlah baris data NISN, Nama, Kelas, dan Jenis Kelamin harus sama!"
      );
      return;
    }

    if (nisnLines.length === 0) {
      alert("⚠️ Tidak ada data yang valid untuk diimport!");
      return;
    }

    if (!confirm(`Akan menambahkan ${nisnLines.length} siswa. Lanjutkan?`)) {
      return;
    }

    setIsBulkSaving(true);

    const studentsData = nisnLines.map((nisn, index) => ({
      nisn: nisn.trim(),
      nama: namaLines[index].trim(),
      kelas: kelasLines[index].trim(),
      jenisKelamin: jenisKelaminLines[index].trim(),
    }));

    fetch(endpoint, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        type: "bulk_siswa",
        students: studentsData,
      }),
    })
      .then(() => {
        alert(
          `✅ Data massal berhasil dikirim! Total: ${studentsData.length} siswa`
        );
        setBulkNisn("");
        setBulkNama("");
        setBulkKelas("");
        setBulkJenisKelamin("");
        setShowBulkImport(false);
        onRefresh();
        setIsBulkSaving(false);
      })
      .catch((error) => {
        console.error("Error:", error);
        alert("❌ Terjadi kesalahan saat import data massal.");
        setIsBulkSaving(false);
      });
  };

  const handleEditStudent = (student: Student) => {
    const newNisn = prompt("Edit NISN:", student.nisn ?? undefined);
    const newName = prompt("Edit nama siswa:", student.name ?? undefined);
    const newClass = prompt("Edit kelas siswa:", student.kelas ?? undefined);
    const jenisKelaminInput = prompt(
      "Edit jenis kelamin (ketik L atau P):",
      student.jenisKelamin ?? undefined
    );

    const newJenisKelamin = jenisKelaminInput?.toUpperCase().trim();
    if (newJenisKelamin && !["L", "P"].includes(newJenisKelamin)) {
      alert("❌ Jenis kelamin harus L atau P!");
      return;
    }

    if (newNisn && newName && newClass && newJenisKelamin) {
      setIsEditing(true);

      fetch(endpoint, {
        method: "POST",
        mode: "no-cors",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          type: "edit",
          nisnLama: student.nisn,
          nisnBaru: newNisn,
          nama: newName,
          kelas: newClass,
          jenisKelamin: newJenisKelamin,
        }),
      })
        .then(() => {
          alert("✅ Data siswa berhasil diperbarui");
          onRefresh();
          setIsEditing(false);
        })
        .catch(() => {
          alert("❌ Gagal memperbarui data");
          setIsEditing(false);
        });
    }
  };

  const handleDeleteStudent = (nisn: string | null | undefined) => {
    if (!nisn) {
      alert("❌ NISN tidak valid untuk penghapusan.");
      return;
    }
    if (confirm("Yakin ingin menghapus siswa ini?")) {
      setIsDeleting(true);

      fetch(endpoint, {
        method: "POST",
        mode: "no-cors",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          type: "delete",
          nisn: nisn,
        }),
      })
        .then(() => {
          alert("🗑️ Data siswa berhasil dihapus");
          onRefresh();
          setIsDeleting(false);
        })
        .catch(() => {
          alert("❌ Gagal menghapus siswa");
          setIsDeleting(false);
        });
    }
  };

  const handleDownloadTemplate = () => {
    const wb = XLSX.utils.book_new();
    const templateData = [
      ["nisn", "nama", "kelas", "jenis kelamin"],
      ["1122", "Andi", "6", "L"],
      ["3171424040", "ALIKA BINTANG SYAM", "6", "L"],
      ["89010", "SENI", "5", "L"],
    ];

    const ws = XLSX.utils.aoa_to_sheet(templateData);
    ws["!cols"] = [{ wch: 15 }, { wch: 30 }, { wch: 10 }, { wch: 15 }];

    for (let row = 2; row <= 1000; row++) {
      const cellRef = `A${row}`;
      if (!ws[cellRef]) ws[cellRef] = { t: "s", v: "" };
      ws[cellRef].z = "@";
    }

    const headerStyle = {
      font: { bold: true, color: { rgb: "FFFFFF" } },
      fill: { fgColor: { rgb: "4472C4" } },
      alignment: { horizontal: "center", vertical: "center" },
    };

    ["A1", "B1", "C1", "D1"].forEach((cell) => {
      if (ws[cell]) {
        ws[cell].s = headerStyle;
      }
    });

    XLSX.utils.book_append_sheet(wb, ws, "Template Data Siswa");
    const date = new Date().toISOString().split("T")[0];
    const fileName = `Template_Data_Siswa_${date}.xlsx`;
    XLSX.writeFile(wb, fileName);

    alert(
      "✅ Template Excel berhasil diunduh!\n\nKolom NISN sudah diformat sebagai TEXT.\nSilakan isi data siswa lalu upload kembali."
    );
  };

  const handleImportExcel = () => {
    fileInputRef.current?.click();
  };

  const handleFileChange = async (
    event: React.ChangeEvent<HTMLInputElement>
  ) => {
    const file = event.target.files?.[0];
    if (!file) return;

    if (!file.name.endsWith(".xlsx") && !file.name.endsWith(".xls")) {
      alert("⚠️ File harus berformat Excel (.xlsx atau .xls)");
      return;
    }

    setIsImporting(true);

    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data);
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
      }) as any[][];

      const headers = jsonData[0];
      const expectedHeaders = ["nisn", "nama", "kelas", "jenis kelamin"];

      const headersMatch = expectedHeaders.every(
        (expected, index) =>
          headers[index]?.toString().toLowerCase().trim() === expected
      );

      if (!headersMatch) {
        alert(
          "⚠️ Format Excel tidak sesuai!\n\nHeader harus: nisn, nama, kelas, jenis kelamin"
        );
        setIsImporting(false);
        return;
      }

      const studentsData: Array<{
        nisn: string;
        nama: string;
        kelas: string;
        jenisKelamin: string;
      }> = [];

      for (let i = 1; i < jsonData.length; i++) {
        const row = jsonData[i];
        if (!row[0] && !row[1] && !row[2] && !row[3]) continue;

        const nisn = String(row[0] || "").trim();
        const nama = String(row[1] || "").trim();
        const kelas = String(row[2] || "").trim();
        let jenisKelamin = String(row[3] || "")
          .trim()
          .toUpperCase();

        if (!nisn || !nama || !kelas || !jenisKelamin) {
          alert(
            `⚠️ Data tidak lengkap pada baris ${
              i + 1
            }!\n\nSemua kolom wajib diisi.`
          );
          setIsImporting(false);
          return;
        }

        if (jenisKelamin === "LAKI-LAKI" || jenisKelamin === "L") {
          jenisKelamin = "L";
        } else if (jenisKelamin === "PEREMPUAN" || jenisKelamin === "P") {
          jenisKelamin = "P";
        } else {
          alert(
            `⚠️ Jenis kelamin tidak valid pada baris ${
              i + 1
            }!\n\nHarus L atau P (Laki-laki atau Perempuan).`
          );
          setIsImporting(false);
          return;
        }

        studentsData.push({ nisn, nama, kelas, jenisKelamin });
      }

      if (studentsData.length === 0) {
        alert("⚠️ Tidak ada data siswa yang valid dalam file Excel!");
        setIsImporting(false);
        return;
      }

      if (
        !confirm(
          `Akan mengimport ${studentsData.length} siswa dari Excel. Lanjutkan?`
        )
      ) {
        setIsImporting(false);
        return;
      }

      await fetch(endpoint, {
        method: "POST",
        mode: "no-cors",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          type: "bulk_siswa",
          students: studentsData,
        }),
      });

      alert(`✅ Berhasil mengimport ${studentsData.length} siswa dari Excel!`);
      onRefresh();

      if (fileInputRef.current) {
        fileInputRef.current.value = "";
      }
    } catch (error) {
      console.error("Error importing Excel:", error);
      alert(
        "❌ Terjadi kesalahan saat mengimport file Excel.\n\nPastikan format file sesuai dan coba lagi."
      );
    } finally {
      setIsImporting(false);
    }
  };

  const handleFotoClick = (student: Student) => {
    setSelectedFotoStudentNisn(student.nisn);
    fotoInputRef.current?.click();
  };

  const handleFotoChange = async (
    event: React.ChangeEvent<HTMLInputElement>
  ) => {
    const file = event.target.files?.[0];
    if (!file || !selectedFotoStudentNisn) return;

    const allowedTypes = ["image/jpeg", "image/jpg", "image/png", "image/webp"];
    if (!allowedTypes.includes(file.type)) {
      alert("⚠️ Format file harus JPG, PNG, atau WEBP!");
      if (fotoInputRef.current) fotoInputRef.current.value = "";
      return;
    }

    setUploadingFotoNisn(selectedFotoStudentNisn);

    try {
      // Kompres gambar
      const compressedBase64 = await compressImage(file, 400, 400);
      const base64Data = compressedBase64.split(",")[1];

      // Gunakan no-cors (tidak bisa baca response)
      await fetch(endpoint, {
        method: "POST",
        mode: "no-cors",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          type: "uploadFoto",
          nisn: selectedFotoStudentNisn,
          fotoBase64: base64Data,
          mimeType: file.type,
          fileName: file.name,
        }),
      });

      // Tunggu sebentar lalu refresh data dari server
      await new Promise((resolve) => setTimeout(resolve, 3000));
      alert("✅ Foto berhasil diupload! Halaman akan diperbarui.");
      onRefresh();
    } catch (error) {
      console.error("Error upload foto:", error);
      alert(
        "❌ Terjadi kesalahan saat upload foto: " +
          (error instanceof Error ? error.message : String(error))
      );
    } finally {
      setUploadingFotoNisn(null);
      setSelectedFotoStudentNisn(null);
      if (fotoInputRef.current) fotoInputRef.current.value = "";
    }
  };

  const compressImage = (
    file: File,
    maxWidth: number,
    maxHeight: number
  ): Promise<string> => {
    return new Promise((resolve, reject) => {
      const canvas = document.createElement("canvas");
      const img = new Image();
      const url = URL.createObjectURL(file);

      img.onload = () => {
        let { width, height } = img;
        if (width > maxWidth || height > maxHeight) {
          const ratio = Math.min(maxWidth / width, maxHeight / height);
          width = Math.round(width * ratio);
          height = Math.round(height * ratio);
        }
        canvas.width = width;
        canvas.height = height;
        const ctx = canvas.getContext("2d")!;
        ctx.drawImage(img, 0, 0, width, height);
        URL.revokeObjectURL(url);
        resolve(canvas.toDataURL("image/jpeg", 0.8));
      };
      img.onerror = reject;
      img.src = url;
    });
  };

  const handleDeleteFoto = async (student: Student) => {
    if (!confirm(`Hapus foto ${student.name}?`)) return;

    setUploadingFotoNisn(student.nisn);
    try {
      await fetch(endpoint, {
        method: "POST",
        mode: "no-cors",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ type: "deleteFoto", nisn: student.nisn }),
      });
      alert("✅ Foto berhasil dihapus!");
      onRefresh();
    } catch {
      alert("❌ Gagal menghapus foto.");
    } finally {
      setUploadingFotoNisn(null);
    }
  };

  // ===== QR CODE FUNCTIONS =====

  const generateQRCode = async (student: Student) => {
    setIsGeneratingQR(true);
    setSelectedStudent(student);
    setShowQRModal(true);

    try {
      // Format QR: NISN:1234567890
      const qrData = `NISN:${student.nisn}`;
      const qrCodeURL = await QRCode.toDataURL(qrData, {
        width: 300,
        margin: 2,
        color: {
          dark: "#000000",
          light: "#FFFFFF",
        },
      });
      setQrCodeDataURL(qrCodeURL);
    } catch (error) {
      console.error("Error generating QR code:", error);
      alert("❌ Gagal generate QR Code");
    } finally {
      setIsGeneratingQR(false);
    }
  };

  const downloadQRCode = () => {
    if (!qrCodeDataURL || !selectedStudent) return;

    const link = document.createElement("a");
    link.href = qrCodeDataURL;
    link.download = `QR_${selectedStudent.name}_${selectedStudent.nisn}.png`;
    link.click();
  };

  const printQRCode = () => {
    if (!qrCodeDataURL || !selectedStudent) return;

    const printWindow = window.open("", "_blank");
    if (!printWindow) return;

    printWindow.document.write(`
      <!DOCTYPE html>
      <html>
        <head>
          <title>Print QR Code - ${selectedStudent.name}</title>
          <style>
            body {
              font-family: Arial, sans-serif;
              display: flex;
              flex-direction: column;
              align-items: center;
              justify-content: center;
              min-height: 100vh;
              margin: 0;
              padding: 20px;
            }
            .qr-container {
              text-align: center;
              border: 2px solid #000;
              padding: 30px;
              border-radius: 10px;
            }
            h2 { margin: 0 0 10px 0; }
            .info { margin: 5px 0; font-size: 14px; }
            img { margin: 20px 0; }
            @media print {
              body { padding: 0; }
            }
          </style>
        </head>
        <body>
          <div class="qr-container">
            <h2>${selectedStudent.name}</h2>
            <div class="info">NISN: ${selectedStudent.nisn}</div>
            <div class="info">Kelas: ${selectedStudent.kelas}</div>
            <img src="${qrCodeDataURL}" alt="QR Code" />
            <div class="info" style="font-size: 12px; color: #666;">
              Scan QR ini untuk absensi
            </div>
          </div>
        </body>
      </html>
    `);

    printWindow.document.close();
    setTimeout(() => {
      printWindow.print();
    }, 250);
  };

  const generateBulkQRCodes = async () => {
    setIsGeneratingBulkQR(true);
    setShowBulkQRModal(true);

    const qrCodes: Array<{ student: Student; qrCode: string }> = [];

    try {
      for (const student of filteredStudents) {
        const qrData = `NISN:${student.nisn}`;
        const qrCodeURL = await QRCode.toDataURL(qrData, {
          width: 200,
          margin: 1,
          color: {
            dark: "#000000",
            light: "#FFFFFF",
          },
        });
        qrCodes.push({ student, qrCode: qrCodeURL });
      }
      setBulkQRCodes(qrCodes);
    } catch (error) {
      console.error("Error generating bulk QR codes:", error);
      alert("❌ Gagal generate QR Code massal");
    } finally {
      setIsGeneratingBulkQR(false);
    }
  };

  const downloadAllQRCodes = async () => {
    if (bulkQRCodes.length === 0) return;

    if (!(window as any).jspdf) {
      await new Promise<void>((resolve, reject) => {
        const script = document.createElement("script");
        script.src =
          "https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js";
        script.onload = () => resolve();
        script.onerror = () => reject(new Error("Failed to load jsPDF"));
        document.head.appendChild(script);
      });
    }

    const { jsPDF } = (window as any).jspdf;

    const pageWidth = 210;
    const pageHeight = 297;
    const cardW = 62;
    const cardH = 90;
    const cols = 3;
    const rows = 3;
    const marginX = (pageWidth - cols * cardW) / (cols + 1);
    const marginY = (pageHeight - rows * cardH) / (rows + 1);

    const doc = new jsPDF({
      orientation: "portrait",
      unit: "mm",
      format: "a4",
    });

    for (let i = 0; i < bulkQRCodes.length; i++) {
      const { student, qrCode } = bulkQRCodes[i];
      const col = i % cols;
      const row = Math.floor(i / cols) % rows;

      if (i > 0 && i % (cols * rows) === 0) {
        doc.addPage();
      }

      const x = marginX + col * (cardW + marginX);
      const y = marginY + row * (cardH + marginY);

      // Card border
      doc.setFillColor(255, 255, 255);
      doc.setDrawColor(100, 116, 139);
      doc.setLineWidth(0.3);
      doc.roundedRect(x, y, cardW, cardH, 3, 3, "FD");

      // Header bar biru
      doc.setFillColor(37, 99, 235);
      doc.roundedRect(x, y, cardW, 14, 3, 3, "F");
      doc.rect(x, y + 10, cardW, 4, "F");

      // Teks header
      doc.setTextColor(255, 255, 255);
      doc.setFontSize(5.5);
      doc.setFont("helvetica", "bold");
      doc.text("KARTU PELAJAR", x + cardW / 2, y + 6, { align: "center" });
      doc.setFontSize(4.5);
      doc.setFont("helvetica", "normal");
      doc.text("ABSENSI DIGITAL", x + cardW / 2, y + 11, { align: "center" });

      // QR Code
      const qrSize = 32;
      const qrX = x + (cardW - qrSize) / 2;
      const qrY = y + 17;
      doc.addImage(qrCode, "PNG", qrX, qrY, qrSize, qrSize);

      // Nama siswa
      doc.setTextColor(30, 41, 59);
      doc.setFontSize(6.5);
      doc.setFont("helvetica", "bold");
      const nameTruncated =
        doc.getTextWidth(student.name || "") > cardW - 6
          ? (student.name || "").substring(0, 20) + "..."
          : student.name || "N/A";
      doc.text(nameTruncated, x + cardW / 2, qrY + qrSize + 5, {
        align: "center",
      });

      // Garis pemisah
      doc.setDrawColor(203, 213, 225);
      doc.setLineWidth(0.2);
      doc.line(x + 4, qrY + qrSize + 7.5, x + cardW - 4, qrY + qrSize + 7.5);

      // Info NISN, Kelas, Jenis Kelamin
      const infoY = qrY + qrSize + 12;
      const labelX = x + 5;
      const valueX = x + 22;

      doc.setFontSize(5.5);
      doc.setFont("helvetica", "normal");
      doc.setTextColor(100, 116, 139);
      doc.text("NISN", labelX, infoY);
      doc.setTextColor(30, 41, 59);
      doc.setFont("helvetica", "bold");
      doc.text(`: ${student.nisn || "N/A"}`, valueX, infoY);

      doc.setFont("helvetica", "normal");
      doc.setTextColor(100, 116, 139);
      doc.text("Kelas", labelX, infoY + 5);
      doc.setTextColor(30, 41, 59);
      doc.setFont("helvetica", "bold");
      doc.text(`: ${student.kelas || "N/A"}`, valueX, infoY + 5);

      doc.setFont("helvetica", "normal");
      doc.setTextColor(100, 116, 139);
      doc.text("J. Kelamin", labelX, infoY + 10);
      doc.setTextColor(30, 41, 59);
      doc.setFont("helvetica", "bold");
      const jkText =
        student.jenisKelamin === "L"
          ? ": Laki-laki"
          : student.jenisKelamin === "P"
          ? ": Perempuan"
          : ": N/A";
      doc.text(jkText, valueX, infoY + 10);

      // Footer
      doc.setFillColor(241, 245, 249);
      doc.rect(x, y + cardH - 10, cardW, 10, "F");
      doc.setDrawColor(100, 116, 239);
      doc.setLineWidth(0.5);
      doc.line(x, y + cardH - 10, x + cardW, y + cardH - 10);
      doc.setFontSize(4);
      doc.setFont("helvetica", "italic");
      doc.setTextColor(148, 163, 184);
      doc.text("Scan QR untuk absensi", x + cardW / 2, y + cardH - 5, {
        align: "center",
      });
    }

    const date = new Date().toISOString().split("T")[0];
    doc.save(`QR_IDCard_Kelas${selectedKelas}_${date}.pdf`);
    alert(`✅ PDF berhasil dibuat! Total: ${bulkQRCodes.length} kartu siswa`);
  };

  const printAllQRCodes = () => {
    if (bulkQRCodes.length === 0) return;

    const printWindow = window.open("", "_blank");
    if (!printWindow) return;

    const qrHTML = bulkQRCodes
      .map(
        (item) => `
      <div class="qr-card">
        <h3>${item.student.name}</h3>
        <div class="info">NISN: ${item.student.nisn}</div>
        <div class="info">Kelas: ${item.student.kelas}</div>
        <img src="${item.qrCode}" alt="QR Code" />
      </div>
    `
      )
      .join("");

    printWindow.document.write(`
      <!DOCTYPE html>
      <html>
        <head>
          <title>Print All QR Codes - Kelas ${selectedKelas}</title>
          <style>
            body {
              font-family: Arial, sans-serif;
              padding: 20px;
            }
            .qr-grid {
              display: grid;
              grid-template-columns: repeat(3, 1fr);
              gap: 20px;
            }
            .qr-card {
              border: 2px solid #000;
              padding: 15px;
              text-align: center;
              border-radius: 8px;
              page-break-inside: avoid;
            }
            h3 {
              margin: 0 0 8px 0;
              font-size: 14px;
            }
            .info {
              font-size: 11px;
              margin: 3px 0;
              color: #333;
            }
            img {
              width: 150px;
              height: 150px;
              margin: 10px 0;
            }
            @media print {
              body { padding: 10px; }
              .qr-card { page-break-inside: avoid; }
            }
          </style>
        </head>
        <body>
          <h1 style="text-align: center;">QR Code Absensi - Kelas ${selectedKelas}</h1>
          <div class="qr-grid">
            ${qrHTML}
          </div>
        </body>
      </html>
    `);

    printWindow.document.close();
    setTimeout(() => {
      printWindow.print();
    }, 500);
  };

  const filteredStudents = React.useMemo(() => {
    if (!searchQuery.trim() && selectedKelas === "Semua") return students;
    const query = searchQuery.toLowerCase().trim();
    return students.filter((student) => {
      const matchesSearchQuery =
        !searchQuery.trim() ||
        (student.name && String(student.name).toLowerCase().includes(query)) ||
        (student.nisn && String(student.nisn).toLowerCase().includes(query));
      const matchesKelas =
        selectedKelas === "Semua" ||
        (student.kelas && String(student.kelas).trim() === selectedKelas);
      return matchesSearchQuery && matchesKelas;
    });
  }, [students, searchQuery, selectedKelas]);

  return (
    <div className="max-w-4xl mx-auto" style={{ paddingBottom: "70px" }}>
      {/* Form Tambah Siswa Tunggal */}
      <div className="bg-white p-6 rounded-lg shadow-md mb-6">
        <h2 className="text-xl font-bold mb-4 text-center text-blue-600">
          Tambah Data Siswa
        </h2>
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-4">
          <input
            type="text"
            placeholder="NISN"
            value={nisn}
            onChange={(e) => setNisn(e.target.value)}
            className="w-full border border-gray-300 px-4 py-2 rounded-lg"
            disabled={isSaving}
          />
          <input
            type="text"
            placeholder="Nama Siswa"
            value={nama}
            onChange={(e) => setNama(e.target.value)}
            className="w-full border border-gray-300 px-4 py-2 rounded-lg"
            disabled={isSaving}
          />
          <input
            type="text"
            placeholder="Kelas"
            value={kelas}
            onChange={(e) => setKelas(e.target.value)}
            className="w-full border border-gray-300 px-4 py-2 rounded-lg"
            disabled={isSaving}
          />
          <select
            value={jenisKelamin}
            onChange={(e) => setJenisKelamin(e.target.value)}
            className="w-full border border-gray-300 px-4 py-2 rounded-lg"
            disabled={isSaving}
          >
            <option value="">Pilih Jenis Kelamin</option>
            <option value="L">L - Laki-laki</option>
            <option value="P">P - Perempuan</option>
          </select>
        </div>
        <div className="grid grid-cols-2 gap-3">
          <button
            onClick={handleSubmit}
            disabled={isSaving || isImporting}
            className={`px-4 py-3 rounded-lg font-medium transition-colors text-sm ${
              isSaving || isImporting
                ? "bg-blue-400 cursor-not-allowed"
                : "bg-blue-600 hover:bg-blue-700"
            } text-white`}
          >
            {isSaving ? "⏳ Menyimpan..." : "➕ Tambah Siswa"}
          </button>
          <button
            onClick={() => setShowBulkImport(!showBulkImport)}
            disabled={isSaving || isImporting}
            className={`px-4 py-3 rounded-lg font-medium transition-colors text-sm ${
              isSaving || isImporting
                ? "bg-green-400 cursor-not-allowed"
                : "bg-green-600 hover:bg-green-700"
            } text-white`}
          >
            📋 Data Massal
          </button>
          <button
            onClick={handleDownloadTemplate}
            disabled={isSaving || isImporting}
            style={{
              backgroundColor: isSaving || isImporting ? "#60a5fa" : "#0ea5e9",
              color: "white",
            }}
            className={`px-4 py-3 rounded-lg font-medium transition-colors text-sm ${
              isSaving || isImporting
                ? "cursor-not-allowed"
                : "hover:brightness-110"
            }`}
          >
            📄 Download Template
          </button>
          <button
            onClick={handleImportExcel}
            disabled={isSaving || isImporting}
            className={`px-4 py-3 rounded-lg font-medium transition-colors text-sm ${
              isSaving || isImporting
                ? "bg-purple-400 cursor-not-allowed"
                : "bg-purple-600 hover:bg-purple-700"
            } text-white`}
          >
            {isImporting ? "⏳ Import..." : "📥 Import Excel"}
          </button>
          <input
            ref={fileInputRef}
            type="file"
            accept=".xlsx,.xls"
            onChange={handleFileChange}
            className="hidden"
          />
          <input
            ref={fotoInputRef}
            type="file"
            accept="image/jpeg,image/jpg,image/png,image/webp"
            onChange={handleFotoChange}
            className="hidden"
          />
        </div>
      </div>

      <div className="bg-purple-50 border border-purple-200 rounded-lg p-4 mb-6">
        <h3 className="text-lg font-semibold text-purple-700 mb-3">
          📋 Cara Import Data dari Excel
        </h3>
        <div className="space-y-3 text-sm text-purple-600">
          <div className="bg-orange-50 border border-orange-300 rounded-lg p-3">
            <p className="font-bold text-orange-700 mb-2">🎯 Langkah Mudah:</p>
            <ol className="list-decimal list-inside space-y-1 text-orange-600">
              <li>
                Klik tombol <strong>"📄 Download Template"</strong> untuk unduh
                template Excel
              </li>
              <li>
                Buka file template, <strong>hapus 3 baris contoh data</strong>
              </li>
              <li>Isi data siswa sesuai format (jangan ubah header!)</li>
              <li>Simpan file Excel</li>
              <li>
                Klik tombol <strong>"📥 Import dari Excel"</strong> dan pilih
                file yang sudah diisi
              </li>
            </ol>
          </div>

          <p className="font-semibold">📄 Format Header Template:</p>
          <div className="bg-white border border-purple-200 rounded p-2 font-mono text-xs">
            nisn | nama | kelas | jenis kelamin
          </div>

          <p className="mt-3">
            <strong>Contoh Data (sudah ada di template):</strong>
          </p>
          <div className="bg-white border border-purple-200 rounded p-2 font-mono text-xs">
            <div>1122 | Andi | 6 | L</div>
            <div>3171424040 | ALIKA BINTANG SYAM | 6 | L</div>
            <div>89010 | SENI | 5 | L</div>
          </div>

          <p className="mt-3">
            <strong>⚠️ Catatan Penting:</strong>
          </p>
          <ul className="list-disc list-inside space-y-1">
            <li>
              <strong>Jangan ubah atau hapus header</strong> (baris 1)
            </li>
            <li>Hapus 3 baris contoh data sebelum mengisi data asli</li>
            <li>
              Jenis kelamin harus: <strong>L</strong> atau <strong>P</strong>
            </li>
            <li>Semua kolom wajib diisi</li>
            <li>
              Format file harus: <strong>.xlsx</strong> atau{" "}
              <strong>.xls</strong>
            </li>
          </ul>
        </div>
      </div>

      {/* Form Bulk Import */}
      {showBulkImport && (
        <div className="bg-white p-6 rounded-lg shadow-md mb-6 border-2 border-green-200">
          <h2 className="text-xl font-bold mb-4 text-center text-green-600">
            Import Data Massal
          </h2>
          <div className="mb-4 p-4 bg-green-50 rounded-lg">
            <p className="text-sm text-green-700 mb-2">
              <strong>Cara penggunaan:</strong>
            </p>
            <p className="text-sm text-green-600">
              1. Copy data dari Excel (pilih kolom NISN, Nama, dan Kelas secara
              terpisah)
              <br />
              2. Paste ke masing-masing kotak di bawah ini
              <br />
              3. Pastikan jumlah baris di setiap kolom sama
              <br />
              4. Klik "Kirim Data Massal"
            </p>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-4">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                NISN (satu per baris)
              </label>
              <textarea
                placeholder="34534534&#10;56565656&#10;12345678"
                value={bulkNisn}
                onChange={(e) => setBulkNisn(e.target.value)}
                className="w-full border border-gray-300 px-4 py-2 rounded-lg h-32 resize-none"
                rows={6}
                disabled={isBulkSaving}
              />
            </div>
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Nama (satu per baris)
              </label>
              <textarea
                placeholder="Andika&#10;Alisa&#10;Budi"
                value={bulkNama}
                onChange={(e) => setBulkNama(e.target.value)}
                className="w-full border border-gray-300 px-4 py-2 rounded-lg h-32 resize-none"
                rows={6}
                disabled={isBulkSaving}
              />
            </div>
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Kelas (satu per baris)
              </label>
              <textarea
                placeholder="3&#10;4&#10;5"
                value={bulkKelas}
                onChange={(e) => setBulkKelas(e.target.value)}
                className="w-full border border-gray-300 px-4 py-2 rounded-lg h-32 resize-none"
                rows={6}
                disabled={isBulkSaving}
              />
            </div>
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Jenis Kelamin (L/P, satu per baris)
              </label>
              <textarea
                placeholder="L&#10;P&#10;L"
                value={bulkJenisKelamin}
                onChange={(e) => setBulkJenisKelamin(e.target.value)}
                className="w-full border border-gray-300 px-4 py-2 rounded-lg h-32 resize-none"
                rows={6}
                disabled={isBulkSaving}
              />
            </div>
          </div>

          <div className="flex justify-center gap-4">
            <button
              onClick={handleBulkImport}
              disabled={isBulkSaving}
              className={`px-6 py-2 rounded-lg font-medium transition-colors ${
                isBulkSaving
                  ? "bg-green-400 cursor-not-allowed"
                  : "bg-green-600 hover:bg-green-700"
              } text-white`}
            >
              {isBulkSaving ? "⏳ Menyimpan..." : "📤 Kirim Data Massal"}
            </button>
            <button
              onClick={() => {
                setBulkNisn("");
                setBulkNama("");
                setBulkKelas("");
                setShowBulkImport(false);
              }}
              disabled={isBulkSaving}
              className={`px-6 py-2 rounded-lg font-medium transition-colors ${
                isBulkSaving
                  ? "bg-gray-400 cursor-not-allowed"
                  : "bg-gray-500 hover:bg-gray-600"
              } text-white`}
            >
              ❌ Batal
            </button>
          </div>
        </div>
      )}

      {/* Pencarian Siswa */}
      <div className="bg-white p-6 rounded-lg shadow-md mb-6">
        <h3 className="text-lg font-semibold text-gray-700 mb-4">
          Pencarian Siswa
        </h3>
        <div className="mb-4">
          <input
            type="text"
            placeholder="Cari berdasarkan nama atau NISN..."
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            className="w-full border border-gray-300 px-4 py-2 rounded-lg"
          />
        </div>
        <div className="mb-4">
          <select
            value={selectedKelas}
            onChange={(e) => setSelectedKelas(e.target.value)}
            className="w-full border border-gray-300 rounded-lg px-4 py-2 shadow-sm bg-white"
          >
            {uniqueClasses.map((kelas) => (
              <option key={kelas} value={kelas}>
                {kelas}
              </option>
            ))}
          </select>
        </div>

        {/* QR Code Bulk Actions */}
        <div className="flex gap-3 justify-center mb-4">
          <button
            onClick={generateBulkQRCodes}
            disabled={filteredStudents.length === 0}
            className={`px-4 py-2 rounded-lg font-medium transition-colors ${
              filteredStudents.length === 0
                ? "bg-gray-300 cursor-not-allowed"
                : "bg-indigo-600 hover:bg-indigo-700"
            } text-white text-sm`}
          >
            🎫 Generate QR Semua ({filteredStudents.length})
          </button>
        </div>
      </div>

      {/* Daftar Siswa */}
      <div className="bg-white p-6 rounded-lg shadow-md">
        <h3 className="text-lg font-semibold text-gray-700 mb-4">
          Daftar Siswa ({filteredStudents.length})
        </h3>
        {loading && students.length === 0 ? (
          <p className="text-center text-gray-500 py-8">
            ⏳ Memuat data siswa...
          </p>
        ) : filteredStudents.length === 0 ? (
          <p className="text-center text-gray-500 py-8">
            {searchQuery || selectedKelas !== "Semua"
              ? "Tidak ada siswa yang cocok dengan pencarian atau filter kelas."
              : "Belum ada data siswa."}
          </p>
        ) : (
          <div className="space-y-3">
            {filteredStudents.map((s) => (
              <div
                key={s.id}
                className="flex justify-between items-center bg-gray-50 border border-gray-200 px-4 py-3 rounded-lg"
              >
                {/* Foto Siswa */}
                <div className="flex items-center gap-3 flex-1">
                  <div className="relative flex-shrink-0">
                    {s.foto ? (
                      <img
                        src={getFotoUrl(s.foto)}
                        alt={s.name}
                        className="w-12 h-12 rounded-full object-cover border-2 border-blue-300 cursor-pointer"
                        onClick={() => {
                          setFotoModalStudent(s);
                          setShowFotoModal(true);
                        }}
                        onError={(e) => {
                          // Coba format URL alternatif jika gagal load
                          const target = e.target as HTMLImageElement;
                          const currentSrc = target.src;

                          if (
                            currentSrc.includes("lh3.googleusercontent.com")
                          ) {
                            // Fallback ke format uc?export=view
                            const fileId = currentSrc.split("/d/")[1];
                            if (fileId) {
                              target.src = `https://drive.google.com/uc?export=view&id=${fileId}`;
                            }
                          } else if (currentSrc.includes("uc?export=view")) {
                            // Fallback ke thumbnail kecil
                            const params = new URLSearchParams(
                              currentSrc.split("?")[1]
                            );
                            const fileId = params.get("id");
                            if (fileId) {
                              target.src = `https://drive.google.com/thumbnail?id=${fileId}&sz=w200`;
                            }
                          } else {
                            // Tampilkan placeholder jika semua gagal
                            target.style.display = "none";
                            target.parentElement!.innerHTML = `<div class="w-12 h-12 rounded-full bg-gray-200 border-2 border-dashed border-gray-400 flex items-center justify-center text-gray-400 text-xl">👤</div>`;
                          }
                        }}
                      />
                    ) : (
                      <div className="w-12 h-12 rounded-full bg-gray-200 border-2 border-dashed border-gray-400 flex items-center justify-center text-gray-400 text-xl">
                        👤
                      </div>
                    )}
                  </div>
                  <div>
                    <p className="font-medium text-gray-800">
                      {s.name || "N/A"}
                    </p>
                    <p className="text-sm text-gray-600">
                      NISN: {s.nisn || "N/A"} | Kelas: {s.kelas || "N/A"} |{" "}
                      {s.jenisKelamin === "L"
                        ? "Laki-laki"
                        : s.jenisKelamin === "P"
                        ? "Perempuan"
                        : s.jenisKelamin || "N/A"}
                    </p>
                  </div>
                </div>
                {/* Tombol Aksi */}
                <div className="flex gap-2 flex-wrap justify-end">
                  <button
                    onClick={() => handleFotoClick(s)}
                    disabled={uploadingFotoNisn === s.nisn}
                    className={`text-xs px-3 py-1 rounded transition-colors ${
                      uploadingFotoNisn === s.nisn
                        ? "bg-gray-300 cursor-not-allowed"
                        : "bg-teal-500 hover:bg-teal-600"
                    } text-white`}
                  >
                    {uploadingFotoNisn === s.nisn ? "⏳" : "📷"}{" "}
                    {s.foto ? "Ganti Foto" : "Upload Foto"}
                  </button>
                  {s.foto && (
                    <button
                      onClick={() => handleDeleteFoto(s)}
                      disabled={uploadingFotoNisn === s.nisn}
                      className="text-xs px-3 py-1 bg-orange-500 hover:bg-orange-600 text-white rounded transition-colors"
                    >
                      🗑️ Hapus Foto
                    </button>
                  )}
                  <button
                    onClick={() => generateQRCode(s)}
                    className="text-xs px-3 py-1 bg-indigo-500 hover:bg-indigo-600 text-white rounded transition-colors"
                  >
                    🎫 QR
                  </button>
                  <button
                    onClick={() => handleEditStudent(s)}
                    disabled={isEditing}
                    className={`text-xs px-3 py-1 rounded transition-colors ${
                      isEditing
                        ? "bg-yellow-400 cursor-not-allowed"
                        : "bg-yellow-500 hover:bg-yellow-600"
                    } text-white`}
                  >
                    {isEditing ? "⏳" : "✏️"} Edit
                  </button>
                  <button
                    onClick={() => handleDeleteStudent(s.nisn)}
                    disabled={isDeleting}
                    className={`text-xs px-3 py-1 rounded transition-colors ${
                      isDeleting
                        ? "bg-red-400 cursor-not-allowed"
                        : "bg-red-500 hover:bg-red-600"
                    } text-white`}
                  >
                    {isDeleting ? "⏳" : "🗑️"} Hapus
                  </button>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>

      {/* Modal QR Code Single */}
      {showQRModal && selectedStudent && (
        <div
          className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50"
          onClick={() => setShowQRModal(false)}
        >
          <div
            className="bg-white rounded-lg shadow-xl p-6 max-w-md w-full mx-4"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex justify-between items-center mb-4">
              <h3 className="text-xl font-bold text-gray-800">
                QR Code Absensi
              </h3>
              <button
                onClick={() => setShowQRModal(false)}
                className="text-gray-400 hover:text-gray-600 text-2xl font-bold"
              >
                ×
              </button>
            </div>

            <div className="text-center mb-4">
              <p className="font-semibold text-lg">{selectedStudent.name}</p>
              <p className="text-sm text-gray-600">
                NISN: {selectedStudent.nisn}
              </p>
              <p className="text-sm text-gray-600">
                Kelas: {selectedStudent.kelas}
              </p>
            </div>

            {isGeneratingQR ? (
              <div className="flex justify-center items-center py-20">
                <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600"></div>
              </div>
            ) : (
              <div className="flex justify-center mb-4">
                <img
                  src={qrCodeDataURL}
                  alt="QR Code"
                  className="border-4 border-gray-300 rounded-lg"
                />
              </div>
            )}

            <div className="flex gap-3 justify-center">
              <button
                onClick={downloadQRCode}
                disabled={isGeneratingQR}
                className="px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-lg font-medium transition-colors"
              >
                📥 Download
              </button>
              <button
                onClick={printQRCode}
                disabled={isGeneratingQR}
                className="px-4 py-2 bg-green-600 hover:bg-green-700 text-white rounded-lg font-medium transition-colors"
              >
                🖨️ Print
              </button>
              <button
                onClick={() => setShowQRModal(false)}
                className="px-4 py-2 bg-gray-500 hover:bg-gray-600 text-white rounded-lg font-medium transition-colors"
              >
                Tutup
              </button>
            </div>
          </div>
        </div>
      )}

      {showFotoModal && fotoModalStudent && (
        <div
          className="fixed inset-0 bg-black bg-opacity-60 flex items-center justify-center z-50"
          onClick={() => setShowFotoModal(false)}
        >
          <div
            className="bg-white rounded-lg shadow-xl p-6 max-w-sm w-full mx-4 text-center"
            onClick={(e) => e.stopPropagation()}
          >
            <h3 className="text-lg font-bold text-gray-800 mb-3">
              {fotoModalStudent.name}
            </h3>
            <img
              src={fotoModalStudent.foto}
              alt={fotoModalStudent.name}
              className="w-48 h-48 rounded-full object-cover mx-auto border-4 border-blue-300 mb-3"
            />
            <p className="text-sm text-gray-600 mb-4">
              NISN: {fotoModalStudent.nisn} | Kelas: {fotoModalStudent.kelas}
            </p>
            <div className="flex gap-3 justify-center">
              <button
                onClick={() => {
                  handleFotoClick(fotoModalStudent);
                  setShowFotoModal(false);
                }}
                className="px-4 py-2 bg-teal-600 hover:bg-teal-700 text-white rounded-lg text-sm font-medium"
              >
                📷 Ganti Foto
              </button>
              <button
                onClick={() => setShowFotoModal(false)}
                className="px-4 py-2 bg-gray-500 hover:bg-gray-600 text-white rounded-lg text-sm font-medium"
              >
                Tutup
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Modal QR Code Bulk */}
      {showBulkQRModal && (
        <div
          className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 overflow-y-auto"
          onClick={() => setShowBulkQRModal(false)}
        >
          <div
            className="bg-white rounded-lg shadow-xl p-6 max-w-6xl w-full mx-4 my-8 max-h-[90vh] overflow-y-auto"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex justify-between items-center mb-4 sticky top-0 bg-white pb-4 border-b">
              <h3 className="text-xl font-bold text-gray-800">
                QR Code Semua Siswa - Kelas {selectedKelas} (
                {bulkQRCodes.length})
              </h3>
              <button
                onClick={() => setShowBulkQRModal(false)}
                className="text-gray-400 hover:text-gray-600 text-2xl font-bold"
              >
                ×
              </button>
            </div>

            {isGeneratingBulkQR ? (
              <div className="flex flex-col justify-center items-center py-20">
                <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mb-4"></div>
                <p className="text-gray-600">Generating QR Codes...</p>
              </div>
            ) : (
              <>
                <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-4 mb-6">
                  {bulkQRCodes.map((item, index) => (
                    <div
                      key={index}
                      className="border-2 border-gray-300 rounded-lg p-3 text-center"
                    >
                      <p className="font-semibold text-sm mb-1">
                        {item.student.name}
                      </p>
                      <p className="text-xs text-gray-600 mb-2">
                        NISN: {item.student.nisn}
                      </p>
                      <img
                        src={item.qrCode}
                        alt="QR Code"
                        className="w-full h-auto mx-auto border border-gray-200 rounded"
                      />
                      <p className="text-xs text-gray-500 mt-1">
                        Kelas {item.student.kelas}
                      </p>
                    </div>
                  ))}
                </div>

                <div className="flex gap-3 justify-center sticky bottom-0 bg-white pt-4 border-t">
                  <button
                    onClick={downloadAllQRCodes}
                    className="px-6 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-lg font-medium transition-colors"
                  >
                    📥 Download Semua
                  </button>
                  <button
                    onClick={printAllQRCodes}
                    className="px-6 py-2 bg-green-600 hover:bg-green-700 text-white rounded-lg font-medium transition-colors"
                  >
                    🖨️ Print Semua
                  </button>
                  <button
                    onClick={() => setShowBulkQRModal(false)}
                    className="px-6 py-2 bg-gray-500 hover:bg-gray-600 text-white rounded-lg font-medium transition-colors"
                  >
                    Tutup
                  </button>
                </div>
              </>
            )}
          </div>
        </div>
      )}
    </div>
  );
};

const AttendanceTab: React.FC<{
  students: Student[];
  onRecapRefresh: () => void;
  statusGuru?: string | null;
  userRole?: "guru" | "siswa" | null;
}> = ({ students, onRecapRefresh, statusGuru, userRole }) => {
  const [attendance, setAttendance] = useState<AttendanceRecord>({});
  const isSiswa = userRole === "siswa";
  const [deletingStudentId, setDeletingStudentId] = useState<string | null>(
    null
  );

  const getLocalDate = () => {
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, "0");
    const day = String(now.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  };

  const [date, setDate] = useState<string>(getLocalDate());
  const [selectedKelas, setSelectedKelas] = useState<string>("Semua");
  const [showDebugInfo, setShowDebugInfo] = useState<boolean>(false);
  const [isSaving, setIsSaving] = useState<boolean>(false);

  // QR Scanner states
  const [scanMode, setScanMode] = useState<boolean>(false);
  const [scannedStudents, setScannedStudents] = useState<Set<string>>(
    new Set()
  );
  const videoRef = useRef<HTMLVideoElement>(null);
  const [scannerError, setScannerError] = useState<string>("");
  const [lastScannedId, setLastScannedId] = useState<string>("");
  const [scannedStudentPhotos, setScannedStudentPhotos] = useState<
    Array<{ student: Student; timestamp: Date; jam: string }>
  >([]);
  const scannerIntervalRef = useRef<number | null>(null);
  const sendingLockRef = useRef<Set<string>>(new Set());
  const [isSendingData, setIsSendingData] = useState<boolean>(false);

  const [tanggalMerahList, setTanggalMerahList] = useState<TanggalMerah[]>([]);
  const [loadingTanggalMerah, setLoadingTanggalMerah] =
    useState<boolean>(false);
  const [jadwalMengajar, setJadwalMengajar] = useState<JadwalMengajar[]>([]);
  const [loadingJadwal, setLoadingJadwal] = useState<boolean>(false);
  const [schoolData, setSchoolData] = useState<SchoolData | null>(null);
  const [existingStudentIds, setExistingStudentIds] = useState<Set<string>>(
    new Set()
  );
  const [isLoadingExistingData, setIsLoadingExistingData] =
    useState<boolean>(false);
  const [existingAttendanceData, setExistingAttendanceData] = useState<any[]>(
    []
  );
  const [fotoAbsensiUrl, setFotoAbsensiUrl] = useState<string | null>(null);
  const [isUploadingFoto, setIsUploadingFoto] = useState<boolean>(false);
  const [previewFoto, setPreviewFoto] = useState<string | null>(null);
  const [fotoToUpload, setFotoToUpload] = useState<string | null>(null);
  const [isDeletingFoto, setIsDeletingFoto] = useState<boolean>(false);
  const [showFotoFullscreen, setShowFotoFullscreen] = useState<boolean>(false);
  const [keterangan, setKeterangan] = useState<{
    [date: string]: { [studentId: string]: string };
  }>({});
  const [showEditModal, setShowEditModal] = useState<boolean>(false);
  const [editData, setEditData] = useState<{
    [studentId: string]: { status: AttendanceStatus; keterangan: string };
  }>({});
  const [isSavingEdit, setIsSavingEdit] = useState<boolean>(false);

  const getFotoUrlAttendance = (fotoUrl: string | undefined): string => {
    if (!fotoUrl) return "";
    if (fotoUrl.includes("lh3.googleusercontent.com")) return fotoUrl;
    if (fotoUrl.includes("uc?export=view&id=")) {
      const fileId = new URLSearchParams(fotoUrl.split("?")[1]).get("id");
      if (fileId) return `https://lh3.googleusercontent.com/d/${fileId}`;
    }
    return fotoUrl;
  };

  const uniqueClasses = React.useMemo(() => {
    const classSet = new Set<string>();
    students.forEach((student) => {
      let kelasValue = student.kelas;
      if (kelasValue != null) {
        kelasValue = String(kelasValue).trim();
        if (
          kelasValue !== "" &&
          kelasValue !== "undefined" &&
          kelasValue !== "null"
        ) {
          classSet.add(kelasValue);
        }
      }
    });
    const classes = Array.from(classSet).sort((a, b) => {
      const aIsNum = /^\d+$/.test(a);
      const bIsNum = /^\d+$/.test(b);
      if (aIsNum && bIsNum) return parseInt(a) - parseInt(b);
      if (aIsNum && !bIsNum) return -1;
      if (!aIsNum && bIsNum) return 1;
      return a.localeCompare(b);
    });
    return ["Semua", ...classes];
  }, [students]);

  const filteredStudents = React.useMemo(() => {
    if (selectedKelas === "Semua") return students;
    return students.filter((student) => {
      if (student.kelas == null) return false;
      const studentKelas = String(student.kelas).trim();
      return studentKelas === selectedKelas;
    });
  }, [students, selectedKelas]);

  const parseTime = (timeStr: string): number => {
    if (!timeStr) return Infinity;
    const [h, m] = timeStr.split(":").map(Number);
    if (isNaN(h) || isNaN(m)) return Infinity;
    return h * 60 + m;
  };

  const sortedStudents = React.useMemo(() => {
    return [...filteredStudents].sort((a, b) => {
      const jamA =
        existingAttendanceData.find((r: any) => r.nama === a.name)?.jam || "";
      const jamB =
        existingAttendanceData.find((r: any) => r.nama === b.name)?.jam || "";

      const hasJamA = jamA !== "" && jamA !== "-";
      const hasJamB = jamB !== "" && jamB !== "-";

      // Siswa dengan jam hadir naik ke atas, diurutkan dari tercepat
      if (hasJamA && hasJamB) return parseTime(jamA) - parseTime(jamB);
      if (hasJamA && !hasJamB) return -1;
      if (!hasJamA && hasJamB) return 1;
      return 0; // keduanya belum hadir, pertahankan urutan asli
    });
  }, [filteredStudents, existingAttendanceData]);

  // ===== QR SCANNER FUNCTIONS =====

  const startScanner = async () => {
    try {
      setScannerError("");
      setScanMode(true);

      const stream = await navigator.mediaDevices.getUserMedia({
        video: { facingMode: "environment" },
      });

      if (videoRef.current) {
        videoRef.current.srcObject = stream;
        await videoRef.current.play();

        scannerIntervalRef.current = window.setInterval(() => {
          if (!scannerIntervalRef.current) return;
          scanQRCode();
        }, 500);
      }
    } catch (error) {
      console.error("Scanner error:", error);
      setScannerError(
        "Tidak dapat mengakses kamera. Pastikan izin kamera telah diberikan."
      );
      stopScanner();
    }
  };

  const scanQRCode = async () => {
    if (!videoRef.current) return;

    try {
      const canvas = document.createElement("canvas");
      const video = videoRef.current;
      canvas.width = video.videoWidth;
      canvas.height = video.videoHeight;
      const ctx = canvas.getContext("2d");
      if (!ctx) return;

      ctx.drawImage(video, 0, 0, canvas.width, canvas.height);
      const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);

      const code = jsQR(imageData.data, imageData.width, imageData.height);

      if (code && code.data) {
        handleScanSuccess(code.data);
      }
    } catch (error) {
      console.error("QR scan error:", error);
    }
  };

  const stopScanner = () => {
    if (scannerIntervalRef.current) {
      clearInterval(scannerIntervalRef.current);
      scannerIntervalRef.current = null;
    }

    if (videoRef.current && videoRef.current.srcObject) {
      const tracks = (videoRef.current.srcObject as MediaStream).getTracks();
      tracks.forEach((track) => track.stop());
      videoRef.current.srcObject = null;
    }

    setScanMode(false);

    // ✅ TAMBAHKAN: Hapus popup dan overlay saat stop scanner
    const popup = document.getElementById("scan-popup");
    const overlay = document.getElementById("scan-overlay");

    if (popup) popup.remove();
    if (overlay) overlay.remove();

    console.log("🛑 Scanner stopped dan popup/overlay dibersihkan");
  };

  const getCurrentTime = () => {
    const now = new Date();
    return `${String(now.getHours()).padStart(2, "0")}:${String(
      now.getMinutes()
    ).padStart(2, "0")}`;
  };

  // ===== FUNGSI KIRIM DATA LANGSUNG =====
  const sendAttendanceData = async (student: Student, jamSekarang: string) => {
    if (sendingLockRef.current.has(student.nisn)) {
      console.log(`🔒 ${student.name} sudah dikunci, skip pengiriman`);
      return;
    }

    sendingLockRef.current.add(student.nisn);
    setIsSendingData(true);

    const formattedDate = formatDateDDMMYYYY(date);

    // ✅ PERBAIKAN: Pastikan keterangan diambil dengan benar
    const keteranganValue =
      (keterangan[date] && keterangan[date][student.id]) || "";

    console.log("📤 Mengirim data:", {
      // TAMBAHKAN LOG
      nama: student.name,
      keterangan: keteranganValue,
    });

    const data = [
      {
        tanggal: formattedDate,
        nama: student.name || "N/A",
        kelas: student.kelas || "N/A",
        nisn: student.nisn || "N/A",
        status: "Hadir",
        keterangan: keteranganValue,
        jam: jamSekarang,
      },
    ];

    try {
      await fetch(endpoint, {
        method: "POST",
        mode: "no-cors",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(data),
      });

      console.log(
        `✅ Data ${student.name} berhasil dikirim dengan keterangan:`,
        keteranganValue
      );
      refreshAttendance();
      onRecapRefresh();
    } catch (error) {
      console.error("❌ Gagal kirim data:", error);
      alert(`❌ Gagal menyimpan data ${student.name}. Silakan coba lagi.`);
      sendingLockRef.current.delete(student.nisn);
      setScannedStudents((prev) => {
        const newSet = new Set(prev);
        newSet.delete(student.id);
        return newSet;
      });
    } finally {
      setIsSendingData(false);
    }
  };

  const handleScanSuccess = async (qrData: string) => {
    try {
      // Hentikan interval PERTAMA KALI sebelum apapun
      if (scannerIntervalRef.current) {
        clearInterval(scannerIntervalRef.current);
        scannerIntervalRef.current = null;
      }

      // Prevent duplicate scans
      if (qrData === lastScannedId) return;
      setLastScannedId(qrData);
      setTimeout(() => setLastScannedId(""), 2000);

      // Extract NISN from QR data (format: NISN:1234567890)
      let nisn = qrData;
      if (qrData.includes("NISN:")) {
        nisn = qrData.split("NISN:")[1];
      }

      // Find student by NISN
      const student = filteredStudents.find(
        (s) => String(s.nisn).trim() === nisn.trim()
      );

      if (!student) {
        playErrorSound();
        alert(`❌ Siswa dengan NISN ${nisn} tidak ditemukan!`);
        return;
      }

      // Check if already scanned
      if (scannedStudents.has(student.id)) {
        playErrorSound();
        alert(`⚠️ ${student.name} sudah diabsen sebelumnya!`);
        return;
      }

      // Check if already has data in database
      if (existingStudentIds.has(student.id)) {
        playErrorSound();
        alert(`⚠️ ${student.name} sudah memiliki data absensi di database!`);
        return;
      }

      // Set attendance locally
      setAttendance((prev) => ({
        ...prev,
        [date]: {
          ...prev[date],
          [student.id]: "Hadir",
        },
      }));

      // Mark as scanned
      setScannedStudents((prev) => new Set([...prev, student.id]));

      // ✅ Hentikan interval + tutup kamera SEGERA setelah scan berhasil
      if (scannerIntervalRef.current) {
        clearInterval(scannerIntervalRef.current);
        scannerIntervalRef.current = null;
      }
      if (videoRef.current && videoRef.current.srcObject) {
        const tracks = (videoRef.current.srcObject as MediaStream).getTracks();
        tracks.forEach((track) => track.stop());
        videoRef.current.srcObject = null;
      }
      setScanMode(false);

      // Tambahkan ke list foto (cegah duplikat)
      const jamScan = getCurrentTime(); // ← Simpan jam di sini

      setScannedStudentPhotos((prev) => {
        const sudahAda = prev.some((item) => item.student.id === student.id);
        if (sudahAda) return prev;
        return [{ student, timestamp: new Date(), jam: jamScan }, ...prev];
      });

      // Tutup kamera
      if (videoRef.current && videoRef.current.srcObject) {
        const tracks = (videoRef.current.srcObject as MediaStream).getTracks();
        tracks.forEach((track) => track.stop());
        videoRef.current.srcObject = null;
      }
      setScanMode(false);

      // Success feedback
      playSuccessSound();
      showSuccessNotification(student.name, student.foto);

      // ✅ KIRIM DATA LANGSUNG KE SERVER
      await sendAttendanceData(student, jamScan);
      // ✅ Stop kamera (interval sudah dihentikan sebelumnya)
    } catch (error) {
      console.error("Scan processing error:", error);
      playErrorSound();
      alert("❌ Error memproses QR Code");
    }
  };

  const handleUploadFotoAbsensi = async () => {
    try {
      // Buat input file yang hanya menerima foto dari kamera
      const input = document.createElement("input");
      input.type = "file";
      input.accept = "image/*";
      input.capture = "environment"; // Langsung buka kamera

      input.onchange = async (e: any) => {
        const file = e.target.files?.[0];
        if (!file) {
          return;
        }

        try {
          // Kompres foto
          const compressedBase64 = await compressImage(file, 2); // Max 2MB

          // Tampilkan preview
          setPreviewFoto(compressedBase64);
          setFotoToUpload(compressedBase64.split(",")[1]); // Simpan base64 tanpa header
        } catch (error) {
          console.error("Error processing foto:", error);
          alert("❌ Gagal memproses foto");
        }
      };

      input.click();
    } catch (error) {
      console.error("Error uploading foto:", error);
      alert("❌ Gagal mengupload foto");
    }
  };

  // Fungsi untuk konfirmasi upload
  const confirmUploadFoto = async () => {
    if (!fotoToUpload) return;

    try {
      setIsUploadingFoto(true);

      const formattedDate = formatDateDDMMYYYY(date);
      const uploadData = {
        type: "uploadFotoAbsensi",
        tanggal: formattedDate,
        kelas: selectedKelas,
        fotoBase64: fotoToUpload,
        mimeType: "image/jpeg",
      };

      const response = await fetch(endpoint, {
        method: "POST",
        mode: "no-cors",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(uploadData),
      });

      console.log("✅ Foto absensi berhasil diupload");
      alert(
        `✅ Foto absensi kelas ${selectedKelas} tanggal ${formattedDate} berhasil disimpan!`
      );

      // Reset preview dan reload foto
      setPreviewFoto(null);
      setFotoToUpload(null);
      syncAll().then(() => loadFotoAbsensi());
    } catch (error) {
      console.error("Error uploading foto:", error);
      alert("❌ Gagal mengupload foto");
    } finally {
      setIsUploadingFoto(false);
    }
  };

  // Fungsi untuk cancel upload
  const cancelUploadFoto = () => {
    setPreviewFoto(null);
    setFotoToUpload(null);
  };

  // Fungsi kompresi gambar
  const compressImage = (file: File, maxSizeMB: number): Promise<string> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();

      reader.onload = (e) => {
        const img = new Image();
        img.onload = () => {
          const canvas = document.createElement("canvas");
          let width = img.width;
          let height = img.height;

          // Hitung ukuran maksimal (misal max 1920px untuk sisi terpanjang)
          const maxDimension = 1920;
          if (width > height && width > maxDimension) {
            height = (height * maxDimension) / width;
            width = maxDimension;
          } else if (height > maxDimension) {
            width = (width * maxDimension) / height;
            height = maxDimension;
          }

          canvas.width = width;
          canvas.height = height;

          const ctx = canvas.getContext("2d");
          if (!ctx) {
            reject(new Error("Canvas context not available"));
            return;
          }

          ctx.drawImage(img, 0, 0, width, height);

          // Kompresi bertahap sampai ukuran di bawah maxSizeMB
          let quality = 0.9;
          let compressedDataUrl = canvas.toDataURL("image/jpeg", quality);

          // Hitung ukuran dalam MB
          const getFileSizeMB = (dataUrl: string) => {
            const base64 = dataUrl.split(",")[1];
            const bytes = atob(base64).length;
            return bytes / (1024 * 1024);
          };

          // Kurangi quality sampai ukuran memenuhi syarat
          while (
            getFileSizeMB(compressedDataUrl) > maxSizeMB &&
            quality > 0.1
          ) {
            quality -= 0.1;
            compressedDataUrl = canvas.toDataURL("image/jpeg", quality);
          }

          console.log(
            `Foto dikompres ke ${getFileSizeMB(compressedDataUrl).toFixed(
              2
            )} MB (quality: ${quality.toFixed(1)})`
          );
          resolve(compressedDataUrl);
        };

        img.onerror = () => reject(new Error("Failed to load image"));
        img.src = e.target?.result as string;
      };

      reader.onerror = () => reject(new Error("Failed to read file"));
      reader.readAsDataURL(file);
    });
  };

  const loadFotoAbsensi = async () => {
    try {
      setFotoAbsensiUrl(
        await getFotoAbsenLocal(formatDateDDMMYYYY(date), selectedKelas)
      );
    } catch (e) {
      console.error(e);
    }
  };

  const handleDeleteFoto = async () => {
    if (
      !confirm(
        `⚠️ Yakin ingin menghapus foto absensi kelas ${selectedKelas} tanggal ${formatDateDDMMYYYY(
          date
        )}?\n\nTindakan ini tidak dapat dibatalkan!`
      )
    ) {
      return;
    }

    setIsDeletingFoto(true);

    try {
      const formattedDate = formatDateDDMMYYYY(date);
      const deleteData = {
        type: "deleteFotoAbsensi",
        tanggal: formattedDate,
        kelas: selectedKelas,
      };

      await fetch(endpoint, {
        method: "POST",
        mode: "no-cors",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(deleteData),
      });

      console.log("✅ Foto absensi berhasil dihapus");
      alert(
        `✅ Foto absensi kelas ${selectedKelas} tanggal ${formattedDate} berhasil dihapus!`
      );

      // Reset state dan reload
      setFotoAbsensiUrl(null);
      syncAll().then(() => loadFotoAbsensi());
    } catch (error) {
      console.error("Error deleting foto:", error);
      alert("❌ Gagal menghapus foto");
    } finally {
      setIsDeletingFoto(false);
    }
  };

  const playSuccessSound = () => {
    const audio = new Audio(
      "data:audio/wav;base64,UklGRnoGAABXQVZFZm10IBAAAAABAAEAQB8AAEAfAAABAAgAZGF0YQoGAACBhYqFbF1fdJivrJBhNjVgodDbq2EcBj+a2/LDciUFLIHO8tiJNwgZaLvt559NEAxQp+PwtmMcBjiR1/LMeSwFJHfH8N2QQAoUXrTp66hVFApGn+DyvmwhBSuBzvLZiTcIGWi77eefTRAMUKfj8LZjHAY4ktfyy3ksBSR3x/DdkEAKFF606+uoVRQKRp/g8r5sIQUrgc7y2Yk3CBlou+3nn00QDFCn4/C2YxwGOJLX8st5LAUkd8fw3ZBACh=="
    );
    audio.play().catch(() => {});
  };

  const playErrorSound = () => {
    const audio = new Audio(
      "data:audio/wav;base64,UklGRiQAAABXQVZFZm10IBAAAAABAAEAESsAACJWAAABAAgAZGF0YQAAAAA="
    );
    audio.play().catch(() => {});
  };

  const showSuccessNotification = (name: string, fotoUrl?: string) => {
    // ✅ PERBAIKAN: Hapus popup dan overlay lama jika masih ada
    const existingPopup = document.getElementById("scan-popup");
    const existingOverlay = document.getElementById("scan-overlay");

    if (existingPopup) {
      existingPopup.remove();
      console.log("🗑️ Popup lama dihapus");
    }

    if (existingOverlay) {
      existingOverlay.remove();
      console.log("🗑️ Overlay lama dihapus");
    }

    const resolvedFoto = fotoUrl
      ? fotoUrl.includes("lh3.googleusercontent.com")
        ? fotoUrl
        : fotoUrl.includes("uc?export=view&id=")
        ? `https://lh3.googleusercontent.com/d/${new URLSearchParams(
            fotoUrl.split("?")[1]
          ).get("id")}`
        : fotoUrl
      : "";

    // ✅ TAMBAHKAN: Buat overlay DULU sebelum popup
    const overlay = document.createElement("div");
    overlay.id = "scan-overlay";
    overlay.style.cssText = `
      position: fixed;
      inset: 0;
      background: rgba(0,0,0,0.5);
      z-index: 9998;
    `;

    // ✅ PERBAIKAN: Fungsi untuk cleanup yang aman
    const cleanup = () => {
      const popup = document.getElementById("scan-popup");
      const overlay = document.getElementById("scan-overlay");

      if (popup) {
        popup.style.opacity = "0";
        popup.style.transition = "opacity 0.2s";
        setTimeout(() => {
          if (popup.parentNode) {
            popup.remove();
          }
        }, 200);
      }

      if (overlay) {
        overlay.style.opacity = "0";
        overlay.style.transition = "opacity 0.2s";
        setTimeout(() => {
          if (overlay.parentNode) {
            overlay.remove();
          }
        }, 200);
      }

      console.log("✅ Cleanup popup dan overlay selesai");
    };

    overlay.addEventListener("click", cleanup);

    const popup = document.createElement("div");
    popup.id = "scan-popup";
    popup.style.cssText = `
      position: fixed;
      top: 50%;
      left: 50%;
      transform: translate(-50%, -50%);
      background: white;
      border-radius: 16px;
      box-shadow: 0 20px 60px rgba(0,0,0,0.4);
      z-index: 9999;
      padding: 24px;
      display: flex;
      flex-direction: column;
      align-items: center;
      gap: 12px;
      border: 4px solid #4ade80;
      min-width: 320px;
      animation: popupFadeIn 0.3s ease;
    `;

    // Tambahkan animasi
    const style = document.createElement("style");
    style.textContent = `
      @keyframes popupFadeIn {
        from { opacity: 0; transform: translate(-50%, -50%) scale(0.8); }
        to { opacity: 1; transform: translate(-50%, -50%) scale(1); }
      }
    `;

    // ✅ PERBAIKAN: Hapus style lama jika ada
    const existingStyle = document.getElementById("popup-animation-style");
    if (existingStyle) {
      existingStyle.remove();
    }
    style.id = "popup-animation-style";
    document.head.appendChild(style);

    popup.innerHTML = `
      <div style="
        background: #4ade80;
        color: white;
        padding: 6px 20px;
        border-radius: 20px;
        font-weight: bold;
        font-size: 16px;
      ">✅ HADIR</div>
  
      ${
        resolvedFoto
          ? `<img
            src="${resolvedFoto}"
            style="
              width: 340px;
              height: 255px;
              object-fit: cover;
              border-radius: 12px;
              border: 3px solid #4ade80;
              box-shadow: 0 4px 20px rgba(0,0,0,0.15);
            "
            onerror="this.style.display='none'; this.nextElementSibling.style.display='flex';"
          />
          <div style="
            display: none;
            width: 340px;
            height: 255px;
            background: #f0fdf4;
            border-radius: 12px;
            border: 3px solid #4ade80;
            align-items: center;
            justify-content: center;
            font-size: 80px;
          ">👤</div>`
          : `<div style="
            display: flex;
            width: 340px;
            height: 255px;
            background: #f0fdf4;
            border-radius: 12px;
            border: 3px solid #4ade80;
            align-items: center;
            justify-content: center;
            font-size: 80px;
          ">👤</div>`
      }
  
      <div style="
        font-size: 22px;
        font-weight: bold;
        color: #1f2937;
        text-align: center;
        max-width: 340px;
      ">${name}</div>
  
      <div style="
        font-size: 13px;
        color: #6b7280;
      ">Tap untuk menutup</div>
    `;

    // ✅ PERBAIKAN: Klik untuk tutup dengan cleanup yang aman
    popup.addEventListener("click", cleanup);

    // ✅ PERBAIKAN: Append overlay DULU, baru popup
    document.body.appendChild(overlay);
    document.body.appendChild(popup);

    // ✅ PERBAIKAN: Auto tutup setelah 4 detik dengan cleanup
    const autoCloseTimer = setTimeout(() => {
      cleanup();
    }, 8000);

    // ✅ TAMBAHKAN: Simpan timer ID untuk bisa di-clear jika ditutup manual
    popup.dataset.timerId = String(autoCloseTimer);

    // ✅ PERBAIKAN: Clear timer jika ditutup manual
    popup.addEventListener("click", () => {
      const timerId = popup.dataset.timerId;
      if (timerId) {
        clearTimeout(Number(timerId));
      }
    });

    console.log("✅ Popup dan overlay berhasil dibuat");
  };

  // Cleanup on unmount
  useEffect(() => {
    return () => {
      stopScanner();
    };
  }, []);

  const loadExistingAttendanceData = async () => {
    setIsLoadingExistingData(true);
    setExistingStudentIds(new Set());
    setExistingAttendanceData([]);
    try {
      const formattedDate = formatDateDDMMYYYY(date);
      const rows = await getAttendanceByDate(formattedDate);
      const filteredData = rows.filter(
        (r: any) =>
          selectedKelas === "Semua" || String(r.kelas).trim() === selectedKelas
      );
      setExistingAttendanceData(filteredData);

      const rec: { [k: string]: AttendanceStatus } = {};
      const ket: { [k: string]: string } = {};
      const ids = new Set<string>();
      const toCheck = selectedKelas === "Semua" ? students : filteredStudents;
      toCheck.forEach((student) => {
        const ex = filteredData.find((r: any) => r.nama === student.name);
        if (ex) {
          rec[student.id] = ex.status as AttendanceStatus;
          ket[student.id] = ex.keterangan || "";
          ids.add(student.id);
        } else {
          // Belum ada data absensi → status dibiarkan kosong
          ket[student.id] = "";
        }
      });
      setAttendance((prev) => ({ ...prev, [date]: rec }));
      setKeterangan((prev) => ({ ...prev, [date]: ket }));
      setExistingStudentIds(ids);
    } catch (error) {
      console.error("Error loading existing attendance data:", error);
    }
    setIsLoadingExistingData(false);
  };

  const refreshAttendance = async () => {
    try {
      await syncAll();
    } catch (e) {
      console.error(e);
    }
    await loadExistingAttendanceData();
  };

  // Auto-sync: ambil perubahan dari aplikasi lain (siswa) setiap 15 detik
  useEffect(() => {
    if (students.length === 0) return;
    const tick = async () => {
      if (
        document.hidden ||
        isSaving ||
        isSavingEdit ||
        showEditModal ||
        scanMode
      )
        return;
      try {
        const changed = await syncAll();
        if (changed.length > 0) await loadExistingAttendanceData();
      } catch (e) {
        console.error("Auto-sync error:", e);
      }
    };
    const id = window.setInterval(tick, 15000);
    document.addEventListener("visibilitychange", tick);
    return () => {
      clearInterval(id);
      document.removeEventListener("visibilitychange", tick);
    };
  }, [
    date,
    selectedKelas,
    students,
    isSaving,
    isSavingEdit,
    showEditModal,
    scanMode,
  ]);

  useEffect(() => {
    if (students.length > 0) {
      loadExistingAttendanceData();
      setScannedStudentPhotos([]);
      setScannedStudents(new Set());
      sendingLockRef.current = new Set(); // ← TAMBAHKAN
    }
  }, [date, selectedKelas, students]);

  useEffect(() => {
    if (students.length && !attendance[date]) {
      const init: { [key: string]: AttendanceStatus } = {};
      setAttendance((prev) => ({ ...prev, [date]: init }));
    }
  }, [date, students, attendance]);

  useEffect(() => {
    if (statusGuru) return; // sudah dimuat induk, tidak perlu fetch DataSekolah lagi
    fetch(`${endpoint}?action=schoolData`)
      .then((res) => (res.ok ? res.json() : Promise.reject()))
      .then((data) => {
        if (data.success && data.data && data.data.length > 0) {
          setSchoolData(data.data[0]);
        }
      })
      .catch(console.error);
  }, []);

  useEffect(() => {
    if (students.length > 0 && selectedKelas !== "Semua") {
      loadFotoAbsensi();
    } else {
      setFotoAbsensiUrl(null);
    }
  }, [date, selectedKelas, students]);

  useEffect(() => {
    // Cleanup function saat component unmount
    return () => {
      // Hapus semua popup dan overlay yang mungkin masih tersisa
      const popup = document.getElementById("scan-popup");
      const overlay = document.getElementById("scan-overlay");
      const style = document.getElementById("popup-animation-style");

      if (popup) popup.remove();
      if (overlay) overlay.remove();
      if (style) style.remove();

      console.log("🧹 Cleanup saat component unmount");
    };
  }, []);

  useEffect(() => {
    return () => {
      document.body.style.overflow = "";
    };
  }, []);

  useEffect(() => {
    fetchTanggalMerah();
    fetchJadwalMengajar();
  }, []);

  const fetchTanggalMerah = async () => {
    try {
      setTanggalMerahList(await getTanggalMerahLocal());
    } catch (e) {
      console.error(e);
    }
  };

  const fetchJadwalMengajar = async () => {
    try {
      setJadwalMengajar(await getJadwalLocal());
    } catch (e) {
      console.error(e);
    }
  };

  const formatDateDDMMYYYY = (isoDate: string): string => {
    const [year, month, day] = isoDate.split("-");
    return `${day}/${month}/${year}`;
  };

  const setStatus = (sid: string, status: AttendanceStatus) => {
    if (isSiswa || existingStudentIds.has(sid)) return;
    setAttendance((prev) => ({
      ...prev,
      [date]: { ...prev[date], [sid]: status },
    }));
  };

  const setKeteranganValue = (sid: string, value: string) => {
    if (existingStudentIds.has(sid)) return;
    setKeterangan((prev) => {
      const currentDateData = prev[date] || {};
      return {
        ...prev,
        [date]: {
          ...currentDateData,
          [sid]: value,
        },
      };
    });
  };

  const handleOpenEditModal = () => {
    const initialEditData: {
      [studentId: string]: { status: AttendanceStatus; keterangan: string };
    } = {};

    filteredStudents.forEach((s) => {
      if (existingStudentIds.has(s.id)) {
        const existingRecord = existingAttendanceData.find(
          (r: any) => r.nama === s.name
        );
        initialEditData[s.id] = {
          status: (existingRecord?.status as AttendanceStatus) || "Hadir",
          keterangan: existingRecord?.keterangan || "",
        };
      }
    });

    setEditData(initialEditData);
    setShowEditModal(true);
    document.body.style.overflow = "hidden"; // ← TAMBAHKAN INI
  };

  const handleSaveEdit = async () => {
    setIsSavingEdit(true);

    // ✅ Hanya kirim siswa yang statusnya BERUBAH dari data asli
    const updates = filteredStudents
      .filter((s) => existingStudentIds.has(s.id) && editData[s.id])
      .filter((s) => {
        const existingRecord = existingAttendanceData.find(
          (r: any) => r.nama === s.name
        );
        const statusBerubah = editData[s.id].status !== existingRecord?.status;
        const keteranganBerubah =
          editData[s.id].keterangan !== (existingRecord?.keterangan || "");
        return statusBerubah || keteranganBerubah; // ← hanya yang berubah
      })
      .map((s) => ({
        tanggal: formatDateDDMMYYYY(date),
        nisn: s.nisn,
        status: editData[s.id].status,
        keterangan: editData[s.id].keterangan,
      }));

    if (updates.length === 0) {
      alert("Tidak ada perubahan yang perlu disimpan.");
      setIsSavingEdit(false);
      return;
    }

    try {
      await fetch(endpoint, {
        method: "POST",
        mode: "no-cors",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ type: "bulkUpdateAttendance", updates }),
      });

      alert("✅ Data kehadiran berhasil diperbarui!");
      setShowEditModal(false);
      document.body.style.overflow = "";
      refreshAttendance();
      onRecapRefresh();
    } catch (error) {
      console.error("Gagal update:", error);
      alert("❌ Gagal memperbarui data. Coba lagi.");
    } finally {
      setIsSavingEdit(false);
    }
  };

  const handleSave = () => {
    if (isSiswa) return;
    setIsSaving(true);
    const formattedDate = formatDateDDMMYYYY(date);
    const studentsToSave = (
      selectedKelas === "Semua" ? students : filteredStudents
    ).filter(
      (s) => !existingStudentIds.has(s.id) && !!attendance[date]?.[s.id]
    );

    if (studentsToSave.length === 0) {
      alert(
        "⚠️ Belum ada status yang dipilih. Klik Hadir / Izin / Sakit / Alpha pada siswa yang ingin disimpan."
      );
      setIsSaving(false);
      return;
    }

    const jamSekarang = getCurrentTime(); // ← TAMBAHKAN DI SINI

    const data = studentsToSave.map((s) => {
      // ✅ PERBAIKAN: Ambil keterangan dengan benar
      const keteranganValue =
        (keterangan[date] && keterangan[date][s.id]) || "";

      console.log("📤 Data untuk", s.name, "keterangan:", keteranganValue); // LOG

      return {
        tanggal: formattedDate,
        nama: s.name || "N/A",
        kelas: s.kelas || "N/A",
        nisn: s.nisn || "N/A",
        status: attendance[date]?.[s.id] || "Hadir",
        keterangan: keteranganValue,
        jam: attendance[date]?.[s.id] === "Hadir" ? jamSekarang : "",
      };
    });

    console.log("📤 Mengirim data ke server:", data); // LOG UNTUK DEBUG

    fetch(endpoint, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(data),
    })
      .then(() => {
        const message =
          selectedKelas === "Semua"
            ? "✅ Data absensi siswa baru semua kelas berhasil dikirim!"
            : `✅ Data absensi siswa baru kelas ${selectedKelas} berhasil dikirim!`;
        alert(message);
        onRecapRefresh();
        refreshAttendance();
        setScannedStudents(new Set());
        setIsSaving(false);
      })
      .catch(() => {
        alert("❌ Gagal kirim data absensi.");
        setIsSaving(false);
      });
  };

  const handleDeleteStudentAttendance = async (student: Student) => {
    if (userRole !== "guru") return;

    const formattedDate = formatDateDDMMYYYY(date);
    if (
      !confirm(
        `Hapus data absensi ${student.name} pada tanggal ${formattedDate}?\n\nStatus, keterangan, dan jam hadir siswa ini akan dihapus. Tindakan ini tidak dapat dibatalkan.`
      )
    ) {
      return;
    }

    setDeletingStudentId(student.id);

    try {
      // status kosong = hapus baris absensi siswa pada tanggal tsb (ditangani bulkUpdateAttendance di GAS)
      await fetch(endpoint, {
        method: "POST",
        mode: "no-cors",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          type: "bulkUpdateAttendance",
          updates: [
            {
              tanggal: formattedDate,
              nisn: student.nisn,
              status: "",
            },
          ],
        }),
      });

      // Bersihkan jejak scan agar siswa ini bisa di-scan ulang
      sendingLockRef.current.delete(student.nisn);
      setScannedStudents((prev) => {
        const newSet = new Set(prev);
        newSet.delete(student.id);
        return newSet;
      });
      setScannedStudentPhotos((prev) =>
        prev.filter((item) => item.student.id !== student.id)
      );

      // Beri waktu Apps Script menyelesaikan penghapusan (no-cors tidak bisa dibaca responsnya)
      await new Promise((resolve) => setTimeout(resolve, 1500));
      await refreshAttendance();
      onRecapRefresh();

      alert(`✅ Data absensi ${student.name} berhasil dihapus.`);
    } catch (error) {
      console.error("Gagal menghapus data absensi siswa:", error);
      alert("❌ Gagal menghapus data absensi. Coba lagi.");
    } finally {
      setDeletingStudentId(null);
    }
  };

  const statusColor: Record<AttendanceStatus, string> = {
    Hadir: "bg-green-500",
    Izin: "bg-yellow-400",
    Sakit: "bg-blue-400",
    Alpha: "bg-red-500",
  };

  const getAttendanceSummary = (): StatusSummary => {
    const summary: StatusSummary = { Hadir: 0, Izin: 0, Sakit: 0, Alpha: 0 };
    filteredStudents.forEach((s) => {
      const status = attendance[date]?.[s.id];
      if (!status) return; // belum ada status → tidak dihitung
      summary[status]++;
    });
    return summary;
  };

  const isSunday = (dateStr: string): boolean => {
    const [year, month, day] = dateStr.split("-");
    const date = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
    return date.getDay() === 0;
  };

  const isInDateRange = (
    currentDate: string,
    startDate: string,
    endDate?: string
  ): boolean => {
    const formatToComparable = (dateStr: string) => {
      const [d, m, y] = dateStr.split("/");
      return new Date(parseInt(y), parseInt(m) - 1, parseInt(d));
    };
    const [year, month, day] = currentDate.split("-");
    const current = new Date(
      parseInt(year),
      parseInt(month) - 1,
      parseInt(day)
    );
    const start = formatToComparable(startDate);
    if (!endDate) return current.getTime() === start.getTime();
    const end = formatToComparable(endDate);
    return current >= start && current <= end;
  };

  const getTanggalMerahInfo = (dateStr: string): TanggalMerah | null => {
    for (const tm of tanggalMerahList) {
      if (isInDateRange(dateStr, tm.tanggal, tm.tanggalAkhir)) return tm;
    }
    return null;
  };

  const isLiburSemester = (dateStr: string): boolean => {
    const info = getTanggalMerahInfo(dateStr);
    if (!info) return false;
    const desc = info.deskripsi.toLowerCase();
    return (
      desc.includes("libur akhir semester") || desc.includes("libur semester")
    );
  };

  const isTanggalMerah = (dateStr: string): boolean => {
    const info = getTanggalMerahInfo(dateStr);
    if (!info) return false;
    const desc = info.deskripsi.toLowerCase();
    return !(
      desc.includes("libur akhir semester") || desc.includes("libur semester")
    );
  };

  const isBukanJadwalMengajar = (dateStr: string): boolean => {
    if ((statusGuru || schoolData?.statusGuru || "Guru Kelas") === "Guru Kelas")
      return false;
    if (selectedKelas === "Semua") return false;
    const jadwal = jadwalMengajar.find((j) => j.kelas === selectedKelas);
    if (!jadwal) return true;
    const [year, month, day] = dateStr.split("-");
    const currentDate = new Date(
      parseInt(year),
      parseInt(month) - 1,
      parseInt(day)
    );
    const dayNames = [
      "Minggu",
      "Senin",
      "Selasa",
      "Rabu",
      "Kamis",
      "Jumat",
      "Sabtu",
    ];
    const dayName = dayNames[currentDate.getDay()];
    const hariJadwal = jadwal.hari
      .split(",")
      .map((h) => h.trim())
      .filter((h) => h.length > 0);
    return !hariJadwal.includes(dayName);
  };

  const attendanceSummary = getAttendanceSummary();
  const allStudentsHaveData =
    existingStudentIds.size === filteredStudents.length;

  if (students.length === 0) {
    return (
      <div className="max-w-4xl mx-auto" style={{ paddingBottom: "70px" }}>
        <div className="bg-white p-6 rounded-lg shadow-md">
          <h2 className="text-2xl font-bold text-center text-blue-700 mb-6">
            📋 Absensi Siswa
          </h2>
          <div className="text-center py-12">
            <div className="bg-yellow-50 border-2 border-yellow-200 rounded-lg p-8 max-w-md mx-auto">
              <div className="text-6xl mb-4">⚠️</div>
              <h3 className="text-2xl font-bold text-yellow-700 mb-2">
                Data Siswa Kosong
              </h3>
              <p className="text-yellow-600 mb-4">
                Anda belum mengisi Data Siswa.
              </p>
              <p className="text-sm text-yellow-500">
                Silakan tambahkan data siswa terlebih dahulu di menu "Data
                Siswa".
              </p>
            </div>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="max-w-4xl mx-auto" style={{ paddingBottom: "70px" }}>
      <div className="bg-white p-6 rounded-lg shadow-md">
        <h2 className="text-2xl font-bold text-center text-blue-700 mb-6">
          📋 Absensi Siswa dengan QR Code
        </h2>

        {isLoadingExistingData && (
          <div className="mb-4 text-center">
            <p className="text-blue-600 text-sm">⏳ Memuat data absensi...</p>
          </div>
        )}

        {allStudentsHaveData && !isLoadingExistingData && (
          <div className="mb-6 bg-green-50 border border-green-200 rounded-lg p-4 text-center">
            <div className="text-green-700 font-semibold text-lg mb-2">
              ✅ Semua siswa sudah diabsen
            </div>
            <p className="text-green-600 text-sm">
              Data absensi untuk tanggal {formatDateDDMMYYYY(date)}
              {selectedKelas !== "Semua" ? ` kelas ${selectedKelas}` : ""} sudah
              lengkap.
            </p>
          </div>
        )}

        {existingStudentIds.size > 0 &&
          !allStudentsHaveData &&
          !isLoadingExistingData && (
            <div className="mb-6 bg-yellow-50 border border-yellow-200 rounded-lg p-4 text-center">
              <div className="text-yellow-700 font-semibold text-lg mb-2">
                ⚠️ Sebagian siswa sudah diabsen
              </div>
              <p className="text-yellow-600 text-sm">
                {existingStudentIds.size} dari {filteredStudents.length} siswa
                sudah memiliki data absensi.
              </p>
            </div>
          )}

        <div className="mb-6 flex flex-col md:flex-row gap-4 items-center justify-center">
          <div className="text-center">
            <p className="text-sm text-gray-500 mb-2">Tanggal</p>
            <input
              type="date"
              value={date}
              onChange={(e) => setDate(e.target.value)}
              className="border border-gray-300 rounded-lg px-1 py-0.5 shadow-sm"
            />
          </div>

          <div className="text-center">
            <p className="text-sm text-gray-500 mb-2">Filter Kelas</p>
            <select
              value={selectedKelas}
              onChange={(e) => setSelectedKelas(e.target.value)}
              className="border border-gray-300 rounded-lg px-1 py-0.5 shadow-sm bg-white min-w-32"
            >
              {uniqueClasses.map((kelas) => (
                <option key={kelas} value={kelas}>
                  {kelas}
                </option>
              ))}
            </select>
          </div>

          {selectedKelas !== "Semua" && (
            <div className="text-center">
              <p className="text-sm text-gray-500 mb-2">Foto Kelas</p>
              <button
                onClick={handleUploadFotoAbsensi}
                disabled={isUploadingFoto || allStudentsHaveData}
                className={`px-4 py-2 rounded-lg font-semibold shadow-md transition-colors ${
                  isUploadingFoto || allStudentsHaveData
                    ? "bg-gray-400 cursor-not-allowed text-white"
                    : "bg-purple-600 hover:bg-purple-700 text-white"
                }`}
              >
                {isUploadingFoto ? "📤 Mengirim..." : "📷 Foto Kelas"}
              </button>
            </div>
          )}

          {/* TAMBAHKAN MODAL PREVIEW FOTO */}
          {previewFoto && (
            <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4">
              <div className="bg-white rounded-lg shadow-xl max-w-2xl w-full max-h-[90vh] overflow-auto">
                <div className="p-6">
                  <h3 className="text-xl font-bold text-gray-800 mb-4 text-center">
                    📷 Preview Foto Absensi
                  </h3>

                  <div className="mb-4 flex justify-center">
                    <img
                      src={previewFoto}
                      alt="Preview Foto Absensi"
                      className="rounded-lg shadow-md max-w-full max-h-96 object-contain"
                    />
                  </div>

                  <div className="bg-blue-50 border border-blue-200 rounded-lg p-3 mb-4">
                    <p className="text-sm text-blue-700 text-center">
                      <strong>Kelas:</strong> {selectedKelas} •{" "}
                      <strong>Tanggal:</strong> {formatDateDDMMYYYY(date)}
                    </p>
                  </div>

                  <div className="flex gap-3 justify-center">
                    <button
                      onClick={cancelUploadFoto}
                      disabled={isUploadingFoto}
                      className="px-6 py-3 bg-gray-500 hover:bg-gray-600 text-white font-semibold rounded-lg shadow-md transition-colors disabled:bg-gray-400 disabled:cursor-not-allowed"
                    >
                      ❌ Batal
                    </button>
                    <button
                      onClick={confirmUploadFoto}
                      disabled={isUploadingFoto}
                      className="px-6 py-3 bg-green-500 hover:bg-green-600 text-white font-semibold rounded-lg shadow-md transition-colors disabled:bg-gray-400 disabled:cursor-not-allowed"
                    >
                      {isUploadingFoto ? "⏳ Mengirim..." : "✅ Kirim Foto"}
                    </button>
                  </div>
                </div>
              </div>
            </div>
          )}

          <div className="text-center">
            <button
              onClick={() => setShowDebugInfo(!showDebugInfo)}
              className="text-sm bg-gray-200 hover:bg-gray-300 px-1 py-0.5 rounded-lg"
            >
              🔍 Info Debug
            </button>
          </div>
        </div>

        {/* Tampilan foto yang sudah diupload */}
        {fotoAbsensiUrl && (
          <div className="mb-6 bg-purple-50 border border-purple-200 rounded-lg p-4">
            <h4 className="font-semibold text-purple-700 mb-2 text-center">
              📷 Foto Absensi Kelas {selectedKelas} - {formatDateDDMMYYYY(date)}
            </h4>
            <div className="flex justify-center mb-3">
              <img
                src={fotoAbsensiUrl}
                alt="Foto Absensi Kelas"
                className="rounded-lg shadow-md max-w-full max-h-96 object-contain cursor-pointer hover:opacity-90 transition-opacity"
                onClick={() => setShowFotoFullscreen(true)}
                title="Klik untuk melihat ukuran penuh"
                onError={(e) => {
                  const target = e.target as HTMLImageElement;
                  target.style.display = "none";
                }}
              />
            </div>

            {/* Tombol Lihat & Hapus */}
            <div className="flex gap-3 justify-center">
              <button
                onClick={() => setShowFotoFullscreen(true)}
                className="px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-lg font-semibold shadow-md transition-colors"
              >
                🔍 Lihat Foto Besar
              </button>
              <button
                onClick={handleDeleteFoto}
                disabled={isDeletingFoto}
                className={`px-4 py-2 rounded-lg font-semibold shadow-md transition-colors ${
                  isDeletingFoto
                    ? "bg-gray-400 cursor-not-allowed text-white"
                    : "bg-red-600 hover:bg-red-700 text-white"
                }`}
              >
                {isDeletingFoto ? "⏳ Menghapus..." : "🗑️ Hapus Foto"}
              </button>
            </div>
          </div>
        )}

        {/* QR Scanner Section */}
        <div className="mb-6">
          <div className="text-center mb-4">
            {!scanMode ? (
              <button
                onClick={startScanner}
                disabled={isSendingData}
                className={`px-6 py-3 rounded-lg font-bold shadow-md transition-colors text-lg ${
                  isSendingData
                    ? "bg-gray-400 cursor-not-allowed"
                    : "bg-blue-600 hover:bg-blue-700"
                } text-white`}
              >
                📷 Mulai Scan QR Code
              </button>
            ) : (
              <button
                onClick={stopScanner}
                disabled={isSendingData}
                className={`px-6 py-3 rounded-lg font-bold shadow-md transition-colors text-lg ${
                  isSendingData
                    ? "bg-gray-400 cursor-not-allowed"
                    : "bg-red-600 hover:bg-red-700"
                } text-white`}
              >
                ⏹️ Hentikan Scanner
              </button>
            )}
          </div>

          {isSendingData && (
            <div className="mb-4 bg-blue-50 border border-blue-200 rounded-lg p-4 text-center">
              <div className="flex items-center justify-center gap-2">
                <svg
                  className="animate-spin h-5 w-5 text-blue-600"
                  xmlns="http://www.w3.org/2000/svg"
                  fill="none"
                  viewBox="0 0 24 24"
                >
                  <circle
                    className="opacity-25"
                    cx="12"
                    cy="12"
                    r="10"
                    stroke="currentColor"
                    strokeWidth="4"
                  ></circle>
                  <path
                    className="opacity-75"
                    fill="currentColor"
                    d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"
                  ></path>
                </svg>
                <span className="text-blue-700 font-semibold">
                  💾 Menyimpan data ke server...
                </span>
              </div>
            </div>
          )}

          {scanMode && (
            <div className="relative border-4 border-blue-500 rounded-lg overflow-hidden">
              <video
                ref={videoRef}
                className="w-full h-96 object-cover bg-black"
                autoPlay
                playsInline
              />
              <div className="absolute top-0 left-0 right-0 bg-blue-600 text-white text-center py-2 font-semibold">
                📱 Arahkan kamera ke QR Code siswa
              </div>
              <div className="absolute bottom-0 left-0 right-0 bg-gradient-to-t from-black to-transparent p-4">
                <p className="text-white text-center text-sm">
                  Sudah Scan: {scannedStudents.size} / {filteredStudents.length}{" "}
                  siswa
                </p>
              </div>
            </div>
          )}

          {scannerError && (
            <div className="mt-4 bg-red-50 border border-red-200 rounded-lg p-4 text-center">
              <p className="text-red-600 font-semibold">❌ {scannerError}</p>
            </div>
          )}

          {scanMode && (
            <div className="mt-4 bg-blue-50 border border-blue-200 rounded-lg p-4">
              <h4 className="font-semibold text-blue-700 mb-2">
                📌 Cara Scan:
              </h4>
              <ol className="list-decimal list-inside text-sm text-blue-600 space-y-1">
                <li>Pastikan QR Code siswa terlihat jelas di kamera</li>
                <li>Tunggu hingga suara "beep" dan notifikasi muncul</li>
                <li>Data akan otomatis tersimpan ke server</li>
                <li>Siswa yang sudah scan tidak bisa scan ulang</li>
              </ol>
            </div>
          )}
        </div>

        {/* Galeri foto siswa yang sudah scan */}
        {scannedStudentPhotos.length > 0 && (
          <div className="mt-4">
            <h4 className="font-semibold text-gray-700 mb-3 text-center">
              ✅ Sudah Hadir ({scannedStudentPhotos.length} siswa)
            </h4>
            <div className="grid grid-cols-3 sm:grid-cols-4 md:grid-cols-5 gap-4">
              {scannedStudentPhotos.map((item, index) => (
                <div key={index} className="flex flex-col items-center gap-1">
                  {item.student.foto ? (
                    <img
                      src={getFotoUrlAttendance(item.student.foto)}
                      alt={item.student.name}
                      className="object-cover shadow-md rounded-lg"
                      style={{
                        width: "113px",
                        height: "85px",
                        border: "3px solid #4ade80",
                        borderRadius: "8px",
                      }}
                      onError={(e) => {
                        const target = e.target as HTMLImageElement;
                        target.style.display = "none";
                        if (target.parentElement) {
                          const div = document.createElement("div");
                          div.style.cssText =
                            "width:113px;height:85px;background:#dcfce7;border:3px solid #4ade80;border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:2rem;";
                          div.innerHTML = "👤";
                          target.parentElement.insertBefore(div, target);
                        }
                      }}
                    />
                  ) : (
                    <div
                      className="flex items-center justify-center shadow-md text-3xl"
                      style={{
                        width: "113px",
                        height: "85px",
                        background: "#dcfce7",
                        border: "3px solid #4ade80",
                        borderRadius: "8px",
                      }}
                    >
                      👤
                    </div>
                  )}
                  <p
                    className="text-xs text-center text-gray-700 font-medium leading-tight truncate"
                    style={{ maxWidth: "113px" }}
                  >
                    {item.student.name.split(" ")[0]}
                  </p>
                  <p className="text-xs text-green-600 font-semibold">
                    {item.jam}
                  </p>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* Modal Fullscreen Foto */}
        {showFotoFullscreen && fotoAbsensiUrl && (
          <div
            className="fixed inset-0 bg-black bg-opacity-90 flex items-center justify-center z-50 p-4"
            onClick={() => setShowFotoFullscreen(false)}
          >
            <div
              className="relative max-w-7xl max-h-full"
              onClick={(e) => e.stopPropagation()}
            >
              {/* Tombol Close */}
              <button
                onClick={() => setShowFotoFullscreen(false)}
                className="absolute top-4 right-4 bg-red-600 hover:bg-red-700 text-white rounded-full w-12 h-12 flex items-center justify-center text-2xl font-bold shadow-lg z-10"
                title="Tutup"
              >
                ×
              </button>

              {/* Info */}
              <div className="absolute top-4 left-4 bg-black bg-opacity-70 text-white px-4 py-2 rounded-lg z-10">
                <p className="font-semibold">
                  Kelas {selectedKelas} • {formatDateDDMMYYYY(date)}
                </p>
              </div>

              {/* Foto Besar */}
              <img
                src={fotoAbsensiUrl}
                alt="Foto Absensi Fullscreen"
                className="rounded-lg shadow-2xl max-w-full max-h-[90vh] object-contain"
                onError={(e) => {
                  const target = e.target as HTMLImageElement;
                  target.style.display = "none";
                  alert("❌ Gagal memuat foto");
                  setShowFotoFullscreen(false);
                }}
              />

              {/* Tombol Download */}
              <div className="absolute bottom-4 left-1/2 transform -translate-x-1/2">
                <a
                  href={fotoAbsensiUrl}
                  download={`Foto_Absensi_${selectedKelas}_${formatDateDDMMYYYY(
                    date
                  )}.jpg`}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="bg-green-600 hover:bg-green-700 text-white px-6 py-3 rounded-lg font-semibold shadow-lg inline-block"
                >
                  📥 Download Foto
                </a>
              </div>
            </div>
          </div>
        )}

        {showEditModal && (
          <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-2">
            <div
              className="bg-white rounded-xl shadow-2xl w-full max-w-2xl flex flex-col"
              style={{ height: "90vh", maxHeight: "90vh" }}
            >
              {/* Header */}
              <div className="flex items-center justify-between px-6 py-4 border-b border-gray-200">
                <div>
                  <h3 className="text-xl font-bold text-gray-800">
                    ✏️ Edit Data Absensi
                  </h3>
                  <p className="text-sm text-gray-500 mt-0.5">
                    {formatDateDDMMYYYY(date)}
                    {selectedKelas !== "Semua"
                      ? ` • Kelas ${selectedKelas}`
                      : ""}
                  </p>
                </div>
                <button
                  onClick={() => {
                    setShowEditModal(false);
                    document.body.style.overflow = "";
                  }}
                  className="text-gray-400 hover:text-gray-600 text-2xl font-bold leading-none"
                >
                  ×
                </button>
              </div>

              {/* Body */}
              <div
                style={{
                  overflowY: "auto",
                  flex: 1,
                  minHeight: 0,
                  padding: "16px 24px",
                }}
              >
                <div style={{ overflowX: "auto" }}>
                  <table
                    style={{
                      minWidth: "480px",
                      width: "100%",
                      borderCollapse: "collapse",
                    }}
                  >
                    <thead>
                      <tr className="bg-gray-50">
                        <th
                          style={{
                            width: "140px",
                            padding: "8px 10px",
                            textAlign: "left",
                            fontSize: "12px",
                            fontWeight: 600,
                            color: "#4b5563",
                            borderBottom: "1px solid #e5e7eb",
                          }}
                        >
                          Nama Siswa
                        </th>
                        <th
                          style={{
                            width: "180px",
                            padding: "8px 10px",
                            textAlign: "center",
                            fontSize: "12px",
                            fontWeight: 600,
                            color: "#4b5563",
                            borderBottom: "1px solid #e5e7eb",
                          }}
                        >
                          Status
                        </th>
                        <th
                          style={{
                            padding: "8px 10px",
                            textAlign: "left",
                            fontSize: "12px",
                            fontWeight: 600,
                            color: "#4b5563",
                            borderBottom: "1px solid #e5e7eb",
                          }}
                        >
                          Keterangan
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredStudents
                        .filter((s) => existingStudentIds.has(s.id))
                        .sort((a, b) => {
                          const jamA =
                            existingAttendanceData.find(
                              (r: any) => r.nama === a.name
                            )?.jam || "";
                          const jamB =
                            existingAttendanceData.find(
                              (r: any) => r.nama === b.name
                            )?.jam || "";
                          return parseTime(jamA) - parseTime(jamB);
                        })
                        .map((s) => {
                          const current = editData[s.id] || {
                            status: "Hadir" as AttendanceStatus,
                            keterangan: "",
                          };
                          const existingRecord = existingAttendanceData.find(
                            (r: any) => r.nama === s.name
                          );
                          return (
                            <tr
                              key={s.id}
                              style={{ borderBottom: "1px solid #f3f4f6" }}
                            >
                              <td
                                style={{
                                  padding: "8px 10px",
                                  verticalAlign: "middle",
                                }}
                              >
                                <p
                                  style={{
                                    fontWeight: 600,
                                    fontSize: "13px",
                                    color: "#1f2937",
                                    margin: 0,
                                  }}
                                >
                                  {s.name}
                                </p>
                                <p
                                  style={{
                                    fontSize: "11px",
                                    color: "#9ca3af",
                                    margin: "2px 0 0",
                                  }}
                                >
                                  Kelas {s.kelas}
                                  {existingRecord?.jam
                                    ? ` • 🕐 ${existingRecord.jam}`
                                    : ""}
                                </p>
                              </td>
                              <td
                                style={{
                                  padding: "8px 10px",
                                  verticalAlign: "middle",
                                }}
                              >
                                <div
                                  style={{
                                    display: "grid",
                                    gridTemplateColumns: "1fr 1fr",
                                    gap: "4px",
                                  }}
                                >
                                  {(
                                    ["Hadir", "Izin", "Sakit", "Alpha"] as const
                                  ).map((st) => (
                                    <button
                                      key={st}
                                      onClick={() =>
                                        setEditData((prev) => ({
                                          ...prev,
                                          [s.id]: { ...current, status: st },
                                        }))
                                      }
                                      style={{
                                        padding: "4px 6px",
                                        borderRadius: "6px",
                                        fontSize: "11px",
                                        fontWeight: 600,
                                        border: "none",
                                        cursor: "pointer",
                                        transition: "opacity 0.15s",
                                        backgroundColor:
                                          current.status === st
                                            ? st === "Hadir"
                                              ? "#22c55e"
                                              : st === "Izin"
                                              ? "#facc15"
                                              : st === "Sakit"
                                              ? "#60a5fa"
                                              : "#f87171"
                                            : "#f3f4f6",
                                        color:
                                          current.status === st
                                            ? st === "Izin"
                                              ? "#713f12"
                                              : "#ffffff"
                                            : "#6b7280",
                                      }}
                                    >
                                      {st}
                                    </button>
                                  ))}
                                </div>
                              </td>
                              <td
                                style={{
                                  padding: "8px 10px",
                                  verticalAlign: "middle",
                                }}
                              >
                                <input
                                  type="text"
                                  value={current.keterangan}
                                  onChange={(e) =>
                                    setEditData((prev) => ({
                                      ...prev,
                                      [s.id]: {
                                        ...current,
                                        keterangan: e.target.value,
                                      },
                                    }))
                                  }
                                  placeholder="Opsional..."
                                  style={{
                                    width: "100%",
                                    padding: "5px 8px",
                                    fontSize: "12px",
                                    border: "1px solid #d1d5db",
                                    borderRadius: "6px",
                                    outline: "none",
                                    boxSizing: "border-box",
                                  }}
                                />
                              </td>
                            </tr>
                          );
                        })}
                    </tbody>
                  </table>
                </div>
              </div>

              {/* Footer */}
              <div className="px-6 py-4 border-t border-gray-200 flex gap-3 justify-end">
                <button
                  onClick={() => {
                    setShowEditModal(false);
                    document.body.style.overflow = "";
                  }}
                  disabled={isSavingEdit}
                  className="px-5 py-2 bg-gray-100 hover:bg-gray-200 text-gray-700 font-semibold rounded-lg transition-colors disabled:opacity-50"
                >
                  Batal
                </button>
                <button
                  onClick={handleSaveEdit}
                  disabled={isSavingEdit}
                  className="px-5 py-2 bg-blue-600 hover:bg-blue-700 text-white font-semibold rounded-lg shadow transition-colors disabled:bg-blue-400 disabled:cursor-not-allowed"
                >
                  {isSavingEdit ? "⏳ Menyimpan..." : "💾 Simpan Perubahan"}
                </button>
              </div>
            </div>
          </div>
        )}

        {showDebugInfo && (
          <div className="mb-6 bg-yellow-50 border border-yellow-200 rounded-lg p-4">
            <h4 className="font-semibold text-yellow-800 mb-2">
              Informasi Debug:
            </h4>
            <div className="text-sm text-yellow-700 space-y-1">
              <p>
                <strong>Total Siswa:</strong> {students.length}
              </p>
              <p>
                <strong>Kelas yang Tersedia:</strong> {uniqueClasses.join(", ")}
              </p>
              <p>
                <strong>Kelas Terpilih:</strong> {selectedKelas}
              </p>
              <p>
                <strong>Siswa Terfilter:</strong> {filteredStudents.length}
              </p>
              <p>
                <strong>Siswa dengan Data Existing:</strong>{" "}
                {existingStudentIds.size}
              </p>
              <p>
                <strong>Siswa Sudah Scan QR:</strong> {scannedStudents.size}
              </p>
              <p>
                <strong>Semua Siswa Sudah Diabsen:</strong>{" "}
                {allStudentsHaveData ? "Ya" : "Tidak"}
              </p>
            </div>
          </div>
        )}

        <div className="mb-4 text-center">
          <p className="text-sm text-gray-600">
            Menampilkan: <span className="font-semibold">{selectedKelas}</span>{" "}
            • Tanggal:{" "}
            <span className="font-semibold">{formatDateDDMMYYYY(date)}</span> •
            Total Siswa:{" "}
            <span className="font-semibold">{filteredStudents.length}</span>
          </p>
        </div>

        {isSunday(date) ? (
          <div className="text-center py-12">
            <div className="bg-red-50 border-2 border-red-200 rounded-lg p-8 max-w-md mx-auto">
              <div className="text-6xl mb-4">🌅</div>
              <h3 className="text-2xl font-bold text-red-700 mb-2">
                Hari Minggu
              </h3>
              <p className="text-red-600">
                Tanggal {formatDateDDMMYYYY(date)} adalah hari Minggu.
              </p>
              <p className="text-sm text-red-500 mt-2">
                Tidak ada kegiatan belajar mengajar.
              </p>
            </div>
          </div>
        ) : isLiburSemester(date) ? (
          <div className="text-center py-12">
            <div className="bg-green-50 border-2 border-green-200 rounded-lg p-8 max-w-md mx-auto">
              <div className="text-6xl mb-4">🏖️</div>
              <h3 className="text-2xl font-bold text-green-700 mb-2">
                Libur Semester
              </h3>
              <p className="text-green-600">
                {getTanggalMerahInfo(date)?.deskripsi}
              </p>
              <p className="text-sm text-green-500 mt-2">
                Tanggal: {formatDateDDMMYYYY(date)}
              </p>
            </div>
          </div>
        ) : isTanggalMerah(date) ? (
          <div className="text-center py-12">
            <div className="bg-yellow-50 border-2 border-yellow-200 rounded-lg p-8 max-w-md mx-auto">
              <div className="text-6xl mb-4">🎉</div>
              <h3 className="text-2xl font-bold text-yellow-700 mb-2">
                Hari Libur Nasional
              </h3>
              <p className="text-yellow-600 font-semibold text-lg">
                {getTanggalMerahInfo(date)?.deskripsi}
              </p>
              <p className="text-sm text-yellow-500 mt-2">
                Tanggal: {formatDateDDMMYYYY(date)}
              </p>
            </div>
          </div>
        ) : isBukanJadwalMengajar(date) ? (
          <div className="text-center py-12">
            <div className="bg-blue-50 border-2 border-blue-200 rounded-lg p-8 max-w-md mx-auto">
              <div className="text-6xl mb-4">📚</div>
              <h3 className="text-2xl font-bold text-blue-700 mb-2">
                Bukan Jadwal Mengajar
              </h3>
              <p className="text-blue-600">
                Hari ini ({formatDateDDMMYYYY(date)}) bukan jadwal mengajar Anda
                untuk kelas {selectedKelas}.
              </p>
              <p className="text-sm text-blue-500 mt-2">
                Silakan pilih tanggal atau kelas lain.
              </p>
            </div>
          </div>
        ) : (
          <>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-6">
              <div className="bg-green-50 border border-green-200 rounded-lg p-3 text-center">
                <div className="text-green-600 font-bold text-lg">
                  {attendanceSummary.Hadir}
                </div>
                <div className="text-green-700 text-sm">Hadir</div>
              </div>
              <div className="bg-yellow-50 border border-yellow-200 rounded-lg p-3 text-center">
                <div className="text-yellow-600 font-bold text-lg">
                  {attendanceSummary.Izin}
                </div>
                <div className="text-yellow-700 text-sm">Izin</div>
              </div>
              <div className="bg-blue-50 border border-blue-200 rounded-lg p-3 text-center">
                <div className="text-blue-600 font-bold text-lg">
                  {attendanceSummary.Sakit}
                </div>
                <div className="text-blue-700 text-sm">Sakit</div>
              </div>
              <div className="bg-red-50 border border-red-200 rounded-lg p-3 text-center">
                <div className="text-red-600 font-bold text-lg">
                  {attendanceSummary.Alpha}
                </div>
                <div className="text-red-700 text-sm">Alpha</div>
              </div>
            </div>

            {userRole === "guru" &&
              existingStudentIds.size > 0 &&
              !isLoadingExistingData && (
                <div className="flex justify-end mb-2">
                  <button
                    onClick={handleOpenEditModal}
                    style={{
                      display: "flex",
                      alignItems: "center",
                      gap: "6px",
                      padding: "8px 14px",
                      backgroundColor: "#f97316",
                      color: "#ffffff",
                      fontWeight: 600,
                      fontSize: "13px",
                      borderRadius: "8px",
                      border: "none",
                      cursor: "pointer",
                      boxShadow: "0 2px 6px rgba(0,0,0,0.15)",
                      whiteSpace: "nowrap", // ← cegah teks wrap/terpotong
                      minWidth: "fit-content", // ← pastikan tombol tidak menyusut
                    }}
                  >
                    ✏️ Edit Data Absensi ({existingStudentIds.size} siswa)
                  </button>
                </div>
              )}

            {isLoadingExistingData && (
              <div className="mb-4 bg-blue-50 border border-blue-200 rounded-lg p-4 text-center">
                <div className="flex items-center justify-center gap-2">
                  <svg
                    className="animate-spin h-5 w-5 text-blue-600"
                    xmlns="http://www.w3.org/2000/svg"
                    fill="none"
                    viewBox="0 0 24 24"
                  >
                    <circle
                      className="opacity-25"
                      cx="12"
                      cy="12"
                      r="10"
                      stroke="currentColor"
                      strokeWidth="4"
                    ></circle>
                    <path
                      className="opacity-75"
                      fill="currentColor"
                      d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"
                    ></path>
                  </svg>
                  <span className="text-blue-700 font-semibold">
                    ⏳ Mohon tunggu, sedang memuat data absensi...
                  </span>
                </div>
              </div>
            )}

            <div className="space-y-4 mb-6 overflow-x-auto">
              <table className="w-full border-collapse">
                <thead>
                  <tr className="bg-gray-100">
                    <th className="border border-gray-200 px-2 py-1 text-center text-sm font-semibold text-gray-700">
                      No.
                    </th>
                    <th className="border border-gray-200 px-2 py-1 text-left text-sm font-semibold text-gray-700">
                      Nama Siswa
                    </th>

                    <th className="border border-gray-200 px-2 py-1 text-center text-sm font-semibold text-gray-700">
                      Status
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {sortedStudents.map((s, index) => {
                    const isExisting = existingStudentIds.has(s.id);
                    const isScanned = scannedStudents.has(s.id);
                    return (
                      <tr
                        key={s.id}
                        className={`border-b border-gray-200 ${
                          isScanned ? "bg-green-50" : ""
                        }`}
                      >
                        <td
                          style={{ width: "1cm" }}
                          className="p-2 text-center"
                        >
                          <span className="text-sm font-medium text-gray-800">
                            {index + 1}
                          </span>
                        </td>
                        <td style={{ width: "5.5cm" }} className="p-2">
                          <p className="text-base font-semibold text-gray-800">
                            {s.name || "N/A"}
                            {isScanned && (
                              <span className="ml-2 text-xs bg-green-500 text-white px-2 py-1 rounded">
                                ✓ Tersimpan
                              </span>
                            )}
                            {isExisting && !isScanned && (
                              <span className="ml-2 text-xs bg-blue-500 text-white px-2 py-1 rounded">
                                ✓ Sudah Ada
                              </span>
                            )}
                          </p>
                          <p className="text-sm text-gray-500">
                            Kelas {s.kelas || "N/A"} • NISN: {s.nisn || "N/A"}
                          </p>
                        </td>

                        <td style={{ width: "5cm" }} className="p-2">
                          <div className="flex flex-col gap-2">
                            {/* Tombol Status */}
                            <div className="flex justify-between">
                              {(
                                ["Hadir", "Izin", "Sakit", "Alpha"] as const
                              ).map((status) => (
                                <button
                                  key={status}
                                  onClick={() => setStatus(s.id, status)}
                                  style={{ width: "1cm" }}
                                  className={`px-1 py-0.5 rounded-lg text-xs font-medium transition-colors ${
                                    isLoadingExistingData
                                      ? "bg-gray-300 text-gray-400 cursor-not-allowed opacity-50"
                                      : attendance[date]?.[s.id] === status
                                      ? `${statusColor[status]} text-white`
                                      : isExisting || isScanned || isSiswa
                                      ? "bg-gray-200 text-gray-500 cursor-not-allowed"
                                      : "bg-white text-gray-700 border border-gray-300 hover:bg-gray-100"
                                  }`}
                                  disabled={
                                    isExisting ||
                                    isLoadingExistingData ||
                                    isScanned ||
                                    isSiswa
                                  }
                                >
                                  {status}
                                </button>
                              ))}
                            </div>

                            {/* TAMBAHKAN: Input Keterangan */}
                            <input
                              type="text"
                              placeholder="Keterangan (opsional)"
                              value={
                                (keterangan[date] && keterangan[date][s.id]) ||
                                ""
                              } // ← PERBAIKAN
                              onChange={(e) =>
                                setKeteranganValue(s.id, e.target.value)
                              }
                              disabled={
                                isExisting ||
                                isLoadingExistingData ||
                                isScanned ||
                                isSiswa
                              }
                              className={`w-full px-2 py-1 text-xs border rounded ${
                                isExisting ||
                                isLoadingExistingData ||
                                isScanned ||
                                isSiswa
                                  ? "bg-gray-100 text-gray-500 cursor-not-allowed"
                                  : "border-gray-300 focus:border-blue-500 focus:outline-none"
                              }`}
                            />
                            {isExisting && (
                              <div className="flex items-center justify-between mt-1">
                                <p className="text-xs text-gray-400">
                                  🕐{" "}
                                  {existingAttendanceData.find(
                                    (r: any) => r.nama === s.name
                                  )?.jam || "-"}
                                </p>
                                {userRole === "guru" && (
                                  <button
                                    onClick={() =>
                                      handleDeleteStudentAttendance(s)
                                    }
                                    disabled={deletingStudentId !== null}
                                    className={`text-xs px-2 py-0.5 rounded text-white transition-colors ${
                                      deletingStudentId !== null
                                        ? "bg-red-300 cursor-not-allowed"
                                        : "bg-red-500 hover:bg-red-600"
                                    }`}
                                  >
                                    {deletingStudentId === s.id ? "⏳" : "🗑️"}{" "}
                                    Hapus
                                  </button>
                                )}
                              </div>
                            )}
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>

            {isSiswa ? null : !allStudentsHaveData && !isLoadingExistingData ? (
              <button
                onClick={handleSave}
                disabled={isSaving}
                className={`w-full py-3 rounded-lg font-bold shadow-md transition-colors ${
                  isSaving
                    ? "bg-blue-400 cursor-not-allowed"
                    : "bg-blue-600 hover:bg-blue-700"
                } text-white`}
              >
                {isSaving
                  ? "⏳ Menyimpan..."
                  : "💾 Simpan Absensi Manual (Siswa Belum Scan)"}
              </button>
            ) : !isLoadingExistingData ? (
              <div className="w-full py-3 rounded-lg font-bold text-center bg-green-100 text-green-700 border border-green-300">
                ✅ Semua siswa sudah diabsen untuk tanggal ini
              </div>
            ) : null}
          </>
        )}
      </div>
    </div>
  );
};

const MonthlyRecapTab: React.FC<{
  onRefresh: () => void;
  uniqueClasses: string[];
  students: Student[];
}> = ({ onRefresh, uniqueClasses, students }) => {
  const [recapData, setRecapData] = useState<MonthlyRecap[]>([]);
  const [selectedKelas, setSelectedKelas] = useState<string>(
    uniqueClasses.find((k) => k !== "Semua") || ""
  );

  const months = [
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
  ] as const;

  // Dapatkan bulan berjalan secara otomatis
  const getCurrentMonth = () => {
    const currentMonthIndex = new Date().getMonth(); // 0-11
    return months[currentMonthIndex];
  };

  const [selectedBulan, setSelectedBulan] = useState<string>(getCurrentMonth());
  const [selectedDate, setSelectedDate] = useState<string>(
    new Date().toISOString().split("T")[0]
  );
  const [loading, setLoading] = useState<boolean>(true);
  const [schoolData, setSchoolData] = useState<SchoolData | null>(null);

  useEffect(() => {
    setLoading(true);
    console.log(
      "Mengambil data rekap dengan kelas:",
      selectedKelas,
      "dan bulan:",
      selectedBulan
    );
    fetch(
      `${endpoint}?action=monthlyRecap&kelas=${selectedKelas}&bulan=${selectedBulan.toLowerCase()}`
    )
      .then((res) => {
        if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
        return res.json();
      })
      .then((data) => {
        console.log("Respons data rekap:", data);
        if (data.success) {
          setRecapData(data.data || []);
        } else {
          alert("❌ Gagal memuat data rekap: " + data.message);
          setRecapData([]);
        }
        setLoading(false);
      })
      .catch((error) => {
        console.error("Error fetch:", error);
        alert("❌ Gagal memuat data rekap. Cek console untuk detail.");
        setRecapData([]);
        setLoading(false);
      });

    // Fetch school data
    fetch(`${endpoint}?action=schoolData`)
      .then((res) => {
        if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
        return res.json();
      })
      .then((data) => {
        if (data.success && data.data && data.data.length > 0) {
          setSchoolData(data.data[0]);
        } else {
          setSchoolData(null);
        }
      })
      .catch((error) => {
        console.error("Error fetching school data:", error);
        alert("❌ Gagal memuat data sekolah. Cek console untuk detail.");
      });
  }, [selectedKelas, selectedBulan, onRefresh]);

  const filteredRecapData = React.useMemo(() => {
    if (!selectedKelas) return recapData;
    return recapData.filter(
      (item) => String(item.kelas).trim() === selectedKelas
    );
  }, [recapData, selectedKelas]);

  const getStatusSummary = (): StatusSummary => {
    const summary: StatusSummary = { Hadir: 0, Izin: 0, Sakit: 0, Alpha: 0 };
    filteredRecapData.forEach((item) => {
      summary.Hadir += item.hadir || 0;
      summary.Alpha += item.alpa || 0;
      summary.Izin += item.izin || 0;
      summary.Sakit += item.sakit || 0;
    });
    return summary;
  };

  const statusSummary = getStatusSummary();

  const downloadExcel = () => {
    const headers = [
      "No.",
      "Nama",
      "Kelas",
      "Hadir",
      "Alpha",
      "Izin",
      "Sakit",
      "% Hadir",
    ];
    const data = [
      headers,
      ...filteredRecapData.map((item, index) => [
        index + 1, // Nomor urut
        item.nama || "N/A",
        item.kelas || "N/A",
        item.hadir || 0,
        item.alpa || 0,
        item.izin || 0,
        item.sakit || 0,
        item.persenHadir !== undefined ? `${item.persenHadir}%` : "N/A",
      ]),
      [
        "",
        "TOTAL",
        "",
        statusSummary.Hadir,
        statusSummary.Alpha,
        statusSummary.Izin,
        statusSummary.Sakit,
        "",
      ],
      [
        "",
        "PERSEN",
        "",
        `${(
          (statusSummary.Hadir /
            (statusSummary.Hadir +
              statusSummary.Alpha +
              statusSummary.Izin +
              statusSummary.Sakit)) *
          100
        ).toFixed(2)}%`,
        `${(
          (statusSummary.Alpha /
            (statusSummary.Hadir +
              statusSummary.Alpha +
              statusSummary.Izin +
              statusSummary.Sakit)) *
          100
        ).toFixed(2)}%`,
        `${(
          (statusSummary.Izin /
            (statusSummary.Hadir +
              statusSummary.Alpha +
              statusSummary.Izin +
              statusSummary.Sakit)) *
          100
        ).toFixed(2)}%`,
        `${(
          (statusSummary.Sakit /
            (statusSummary.Hadir +
              statusSummary.Alpha +
              statusSummary.Izin +
              statusSummary.Sakit)) *
          100
        ).toFixed(2)}%`,
        "",
      ],
    ];

    const ws = XLSX.utils.aoa_to_sheet(data);
    ws["!cols"] = [
      { wch: 5 }, // Lebar kolom No. (sempit)
      { wch: 25 }, // Nama
      { wch: 10 }, // Kelas
      { wch: 10 }, // Hadir
      { wch: 10 }, // Alpha
      { wch: 10 }, // Izin
      { wch: 10 }, // Sakit
      { wch: 10 }, // % Hadir
    ];
    const headerStyle = {
      font: { bold: true },
      fill: { fgColor: { rgb: "FFFF00" } },
      alignment: { horizontal: "center" },
    };
    const totalStyle = {
      font: { bold: true },
      fill: { fgColor: { rgb: "D3D3D3" } },
      alignment: { horizontal: "center" },
    };
    const percentStyle = {
      font: { bold: true },
      fill: { fgColor: { rgb: "D3D3D3" } },
      alignment: { horizontal: "center" },
    };
    headers.forEach((header, index) => {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: index });
      ws[cellAddress] = { ...ws[cellAddress], s: headerStyle };
    });
    const totalRow = filteredRecapData.length + 1;
    ["A", "B", "C", "D", "E", "F", "G", "H"].forEach((col, idx) => {
      const cellAddress = `${col}${totalRow}`;
      ws[cellAddress] = { ...ws[cellAddress], s: totalStyle };
    });
    const percentRow = filteredRecapData.length + 2;
    ["A", "B", "C", "D", "E", "F", "G", "H"].forEach((col, idx) => {
      const cellAddress = `${col}${percentRow}`;
      ws[cellAddress] = { ...ws[cellAddress], s: percentStyle };
    });

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Rekap Bulanan");

    const date = new Date()
      .toLocaleString("id-ID", {
        day: "2-digit",
        month: "long",
        year: "numeric",
        hour: "2-digit",
        minute: "2-digit",
        hour12: false,
      })
      .replace(/ /g, "_")
      .replace(/:/g, "-");
    const fileName = `Rekap_Bulanan_${selectedBulan}_${selectedKelas}_${date}.xlsx`;
    XLSX.writeFile(wb, fileName);
  };

  const downloadPDF = async () => {
    const doc = new jsPDF();
    const pageWidth = doc.internal.pageSize.getWidth();
    const margin = 14;
    const lineSpacing = 5;
    let currentY = margin;

    doc.setFont("Times", "roman");

    // Title - Format sama dengan Daftar Hadir
    const namaSekolah = schoolData?.namaSekolah || "UPT SDN 13 BATANG";

    // ✅ TAMBAHKAN: Ambil tahun dari selectedDate
    const tahunDariTanggal = new Date(selectedDate).getFullYear();

    // Judul dalam 1 baris - gunakan tahunDariTanggal
    const title = `REKAP ABSENSI SISWA KELAS ${selectedKelas}  ${namaSekolah}  ${selectedBulan.toUpperCase()} ${tahunDariTanggal}`;

    doc.setFontSize(12); // Ukuran font sama dengan daftar hadir
    doc.setFont("Times", "bold");
    doc.text(title, pageWidth / 2, currentY, { align: "center" });

    currentY += 10;

    // Table headers and data
    const headers = [
      "No.",
      "Nama",
      "Kelas",
      "Hadir",
      "Alpha",
      "Izin",
      "Sakit",
      "% Hadir",
    ];
    const body = filteredRecapData.map((item, index) => [
      index + 1, // Nomor urut
      item.nama || "N/A",
      item.kelas || "N/A",
      item.hadir || 0,
      item.alpa || 0,
      item.izin || 0,
      item.sakit || 0,
      item.persenHadir !== undefined ? `${item.persenHadir}%` : "N/A",
    ]);

    const totalRow = [
      "",
      "TOTAL",
      "",
      statusSummary.Hadir,
      statusSummary.Alpha,
      statusSummary.Izin,
      statusSummary.Sakit,
      "",
    ];

    const percentRow = [
      "",
      "PERSEN",
      "",
      `${(
        (statusSummary.Hadir /
          (statusSummary.Hadir +
            statusSummary.Alpha +
            statusSummary.Izin +
            statusSummary.Sakit)) *
        100
      ).toFixed(2)}%`,
      `${(
        (statusSummary.Alpha /
          (statusSummary.Hadir +
            statusSummary.Alpha +
            statusSummary.Izin +
            statusSummary.Sakit)) *
        100
      ).toFixed(2)}%`,
      `${(
        (statusSummary.Izin /
          (statusSummary.Hadir +
            statusSummary.Alpha +
            statusSummary.Izin +
            statusSummary.Sakit)) *
        100
      ).toFixed(2)}%`,
      `${(
        (statusSummary.Sakit /
          (statusSummary.Hadir +
            statusSummary.Alpha +
            statusSummary.Izin +
            statusSummary.Sakit)) *
        100
      ).toFixed(2)}%`,
      "",
    ];

    autoTable(doc, {
      head: [headers],
      body: [...body, totalRow, percentRow],
      startY: currentY,
      styles: { font: "Times", fontSize: 8, cellPadding: 2 },
      headStyles: {
        fillColor: [255, 255, 0],
        textColor: [0, 0, 0],
        fontStyle: "bold",
      },
      alternateRowStyles: { fillColor: [240, 240, 240] },
      columnStyles: {
        0: { cellWidth: 10 }, // No. (sempit)
        1: { cellWidth: 50 }, // Nama
        2: { cellWidth: 20 }, // Kelas
        3: { cellWidth: 20 }, // Hadir
        4: { cellWidth: 20 }, // Alpha
        5: { cellWidth: 20 }, // Izin
        6: { cellWidth: 20 }, // Sakit
        7: { cellWidth: 20 }, // % Hadir
      },
    });

    // Update currentY after the table
    currentY = (doc as any).lastAutoTable.finalY + 10;

    // ✅ TAMBAHAN BARU: Tabel Informasi Jumlah Siswa
    const pageHeight = doc.internal.pageSize.getHeight();
    const bottomMargin = 20;
    const spaceNeededForStudentTable = 20;
    const spaceNeededForSignatures = 60; // Ruang untuk tanda tangan

    // ✅ CEK APAKAH TANDA TANGAN + TABEL SISWA MUAT DI HALAMAN INI
    // Cek apakah ada cukup ruang untuk tabel jumlah siswa DAN tanda tangan
    if (
      currentY + spaceNeededForStudentTable + spaceNeededForSignatures >
      pageHeight - bottomMargin
    ) {
      doc.addPage();
      currentY = margin;
    }

    // Hitung jumlah siswa berdasarkan jenis kelamin
    const genderSummary = filteredRecapData.reduce(
      (acc, student) => {
        // Ambil data siswa dari filteredRecapData untuk mendapatkan jenis kelamin
        const studentData = students.find((s) => s.name === student.nama);
        if (studentData) {
          const jenisKelamin = String(studentData.jenisKelamin || "")
            .trim()
            .toUpperCase();
          if (jenisKelamin === "L" || jenisKelamin === "LAKI-LAKI") {
            acc.lakiLaki++;
          } else if (jenisKelamin === "P" || jenisKelamin === "PEREMPUAN") {
            acc.perempuan++;
          }
        }
        return acc;
      },
      { lakiLaki: 0, perempuan: 0 }
    );

    const totalSiswa = genderSummary.lakiLaki + genderSummary.perempuan;

    doc.setFontSize(10);
    doc.setFont("Times", "bold");
    doc.text("JUMLAH SISWA", margin, currentY, { align: "left" });
    currentY += 3;

    const tableWidth = (pageWidth - 2 * margin) * 0.4;

    autoTable(doc, {
      head: [["LAKI-LAKI", "PEREMPUAN", "TOTAL SISWA"]],
      body: [
        [
          genderSummary.lakiLaki.toString(),
          genderSummary.perempuan.toString(),
          totalSiswa.toString(),
        ],
      ],
      startY: currentY,
      margin: { left: margin, right: pageWidth - margin - tableWidth },
      tableWidth: tableWidth,
      theme: "grid",
      styles: {
        font: "Times",
        fontSize: 7,
        cellPadding: 1,
        halign: "center",
        valign: "middle",
        lineWidth: 0.5,
      },
      headStyles: {
        fillColor: [255, 255, 255],
        textColor: [0, 0, 0],
        fontStyle: "bold",
        lineWidth: 1,
      },
      bodyStyles: {
        fillColor: [255, 255, 255],
        textColor: [0, 0, 0],
        fontStyle: "bold",
        fontSize: 10,
        lineWidth: 1,
      },
      columnStyles: {
        0: { cellWidth: tableWidth / 3, fillColor: [255, 255, 255] },
        1: { cellWidth: tableWidth / 3, fillColor: [255, 255, 255] },
        2: { cellWidth: tableWidth / 3, fillColor: [255, 255, 255] },
      },
    });

    currentY = (doc as any).lastAutoTable.finalY + 10;

    // Add school data (Principal and Teacher details)
    if (schoolData) {
      doc.setFontSize(10);
      doc.setFont("Times", "roman");

      // Add place and date above Guru Kelas, centered
      const formattedDate = new Date(selectedDate).toLocaleDateString("id-ID", {
        day: "2-digit",
        month: "long",
        year: "numeric",
      });
      const placeDateText = `${
        schoolData.namaKota || "Makassar"
      }, ${formattedDate}`;
      const rightColumnX = pageWidth - margin - 50; // Signature width is 50
      doc.text(placeDateText, rightColumnX + 25, currentY - 1, {
        align: "center",
      });
      currentY += 5; // Keep close to "Guru Kelas"

      // Principal Section
      const principalText = [
        "Kepala Sekolah,",
        "",
        "",
        `( ${schoolData.namaKepsek || "N/A"} )`,
        `NIP: ${schoolData.nipKepsek || "N/A"}`,
      ];
      const teacherText = [
        "Guru Kelas,",
        "",
        "",
        `( ${schoolData.namaGuru || "N/A"} )`,
        `NIP: ${schoolData.nipGuru || "N/A"}`,
      ];

      // Calculate width for signatures
      const signatureWidth = 30;
      const signatureHeight = 20;
      const leftColumnX = margin;

      // Principal signature and text
      if (schoolData.ttdKepsek) {
        try {
          const canvas = document.createElement("canvas");
          canvas.width = 150; // Sesuaikan ukuran canvas (lebar lebih besar untuk tanda tangan panjang)
          canvas.height = 50; // Sesuaikan ukuran canvas (tinggi cukup untuk garis tanda tangan)
          const ctx = canvas.getContext("2d");
          const v = await Canvg.from(ctx, schoolData.ttdKepsek); // schoolData.ttdKepsek adalah base64 SVG
          v.start();
          const pngData = canvas.toDataURL("image/png");
          doc.addImage(
            pngData,
            "PNG",
            leftColumnX + 10,
            currentY - 3,
            signatureWidth,
            signatureHeight
          ); // Sesuaikan posisi sesuai asli
        } catch (error) {
          console.error("Error rendering Kepsek signature:", error);
          doc.setFontSize(10);
          doc.text(
            "Gagal render tanda tangan Kepala Sekolah.",
            leftColumnX + 10,
            currentY - 3 + 10
          );
        }
      }

      // Pisahkan "Kepala Sekolah" dengan posisi yang lebih tinggi
      doc.text("Kepala Sekolah,", leftColumnX + 25, currentY - 2, {
        align: "center",
      });

      // Kosong dan kosong
      doc.text("", leftColumnX + 25, currentY + lineSpacing, {
        align: "center",
      });
      doc.text("", leftColumnX + 25, currentY + 2 * lineSpacing, {
        align: "center",
      });

      // Nama kepala sekolah dengan format bold dan underline
      const principalName = schoolData.namaKepsek || "N/A";
      doc.setFont("Times", "bold");
      doc.text(principalName, leftColumnX + 25, currentY + 3.5 * lineSpacing, {
        align: "center",
      });

      // Add underline to principal name
      const principalNameText = principalName;
      const textWidth = doc.getTextWidth(principalNameText);
      const textX = leftColumnX + 25 - textWidth / 2;
      doc.line(
        textX,
        currentY + 3.5 * lineSpacing + 1,
        textX + textWidth,
        currentY + 3.5 * lineSpacing + 1
      );

      // Reset font and add NIP
      doc.setFont("Times", "roman");
      doc.text(
        `NIP. ${schoolData.nipKepsek || "N/A"}`,
        leftColumnX + 25,
        currentY + 4.5 * lineSpacing,
        {
          align: "center",
        }
      );

      // Teacher signature and text
      if (schoolData.ttdGuru) {
        try {
          const canvas = document.createElement("canvas");
          canvas.width = 150; // Sesuaikan ukuran canvas
          canvas.height = 50;
          const ctx = canvas.getContext("2d");
          const v = await Canvg.from(ctx, schoolData.ttdGuru); // schoolData.ttdGuru adalah base64 SVG
          v.start();
          const pngData = canvas.toDataURL("image/png");
          doc.addImage(
            pngData,
            "PNG",
            rightColumnX + 10,
            currentY - 5,
            signatureWidth,
            signatureHeight
          ); // Sesuaikan posisi sesuai asli
        } catch (error) {
          console.error("Error rendering Guru signature:", error);
          doc.setFontSize(10);
          doc.text(
            "Gagal render tanda tangan Guru.",
            rightColumnX + 10,
            currentY - 5 + 10
          );
        }
      }

      // Pisahkan "Guru Kelas" dengan posisi yang lebih tinggi
      doc.text(
        `${schoolData.statusGuru || "Guru Kelas"},`,
        rightColumnX + 25,
        currentY - 2,
        {
          align: "center",
        }
      );

      // Kosong dan kosong
      doc.text("", rightColumnX + 25, currentY + lineSpacing, {
        align: "center",
      });
      doc.text("", rightColumnX + 25, currentY + 2 * lineSpacing, {
        align: "center",
      });

      // Nama guru dengan format bold dan underline
      const teacherName = schoolData.namaGuru || "N/A";
      doc.setFont("Times", "bold");
      doc.text(teacherName, rightColumnX + 25, currentY + 3.5 * lineSpacing, {
        align: "center",
      });

      // Add underline to teacher name
      const teacherNameText = teacherName;
      const teacherTextWidth = doc.getTextWidth(teacherNameText);
      const teacherTextX = rightColumnX + 25 - teacherTextWidth / 2;
      doc.line(
        teacherTextX,
        currentY + 3.5 * lineSpacing + 1,
        teacherTextX + teacherTextWidth,
        currentY + 3.5 * lineSpacing + 1
      );

      // Reset font and add NIP
      doc.setFont("Times", "roman");
      doc.text(
        `NIP. ${schoolData.nipGuru || "N/A"}`,
        rightColumnX + 25,
        currentY + 4.5 * lineSpacing,
        {
          align: "center",
        }
      );
    } else {
      doc.setFontSize(10);
      doc.text("Data sekolah tidak tersedia.", margin, currentY);
    }

    const date = new Date()
      .toLocaleString("id-ID", {
        day: "2-digit",
        month: "long",
        year: "numeric",
        hour: "2-digit",
        minute: "2-digit",
        hour12: false,
      })
      .replace(/ /g, "_")
      .replace(/:/g, "-");
    const fileName = `Rekap_Bulanan_${selectedBulan}_${selectedKelas}_${date}.pdf`;
    doc.save(fileName);
  };

  return (
    <div className="max-w-4xl mx-auto" style={{ paddingBottom: "70px" }}>
      <div className="bg-white p-6 rounded-lg shadow-md">
        <h2 className="text-2xl font-bold text-center text-blue-700 mb-6">
          📊 Rekap Absensi Bulanan
        </h2>
        <div className="mb-6 flex flex-col md:flex-row gap-4 items-center justify-center">
          <div className="text-center">
            <p className="text-sm text-gray-500 mb-2">Filter Kelas</p>
            <select
              value={selectedKelas}
              onChange={(e) => {
                console.log("Mengubah filter kelas ke:", e.target.value);
                setSelectedKelas(e.target.value);
              }}
              className="border border-gray-300 rounded-lg px-1 py-0.5 shadow-sm bg-white min-w-32"
            >
              {uniqueClasses
                .filter((kelas) => kelas !== "Semua")
                .map((kelas) => (
                  <option key={kelas} value={kelas}>
                    {kelas}
                  </option>
                ))}
            </select>
          </div>
          <div className="text-center">
            <p className="text-sm text-gray-500 mb-2">Pilih Bulan</p>
            <select
              value={selectedBulan}
              onChange={(e) => setSelectedBulan(e.target.value)}
              className="border border-gray-300 rounded-lg px-1 py-0.5 shadow-sm bg-white min-w-32"
            >
              {months.map((month) => (
                <option key={month} value={month}>
                  {month}
                </option>
              ))}
            </select>
          </div>
        </div>

        {/* Separator line and PDF settings section */}
        <div className="border-t border-gray-200 pt-4 mb-6">
          <p className="text-center text-sm font-medium text-gray-700 mb-4">
            Pengaturan Tanggal & Nama Tempat <br /> untuk Rekap Bulanan pada
            File PDF
          </p>
          <div className="flex flex-col md:flex-row gap-4 items-center justify-center">
            <div className="text-center">
              <p className="text-sm text-gray-500 mb-2">Pilih Tanggal</p>
              <input
                type="date"
                value={selectedDate}
                onChange={(e) => setSelectedDate(e.target.value)}
                className="border border-gray-300 rounded-lg px-1 py-0.5 shadow-sm bg-white min-w-32"
              />
            </div>
          </div>
        </div>

        <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-6">
          <div className="bg-green-50 border border-green-200 rounded-lg p-3 text-center">
            <div className="text-green-600 font-bold text-lg">
              {statusSummary.Hadir}
            </div>
            <div className="text-green-700 text-sm">Hadir</div>
          </div>
          <div className="bg-yellow-50 border border-yellow-200 rounded-lg p-3 text-center">
            <div className="text-yellow-600 font-bold text-lg">
              {statusSummary.Izin}
            </div>
            <div className="text-yellow-700 text-sm">Izin</div>
          </div>
          <div className="bg-blue-50 border border-blue-200 rounded-lg p-3 text-center">
            <div className="text-blue-600 font-bold text-lg">
              {statusSummary.Sakit}
            </div>
            <div className="text-blue-700 text-sm">Sakit</div>
          </div>
          <div className="bg-red-50 border border-red-200 rounded-lg p-3 text-center">
            <div className="text-red-600 font-bold text-lg">
              {statusSummary.Alpha}
            </div>
            <div className="text-red-700 text-sm">Alpha</div>
          </div>
        </div>

        {loading ? (
          <div className="text-center py-8">
            <p className="text-gray-500">Memuat rekap...</p>
          </div>
        ) : filteredRecapData.length === 0 ? (
          <div className="text-center py-8">
            <p className="text-gray-500">
              Tidak ada data rekap untuk {selectedBulan} kelas {selectedKelas}.
            </p>
            <p className="text-sm text-gray-400 mt-2">
              Coba pilih kelas atau bulan lain.
            </p>
          </div>
        ) : (
          <>
            {(() => {
              const total =
                statusSummary.Hadir +
                statusSummary.Alpha +
                statusSummary.Izin +
                statusSummary.Sakit;
              const pctHadir =
                total > 0
                  ? ((statusSummary.Hadir / total) * 100).toFixed(1)
                  : "0";
              const pctAlpha =
                total > 0
                  ? ((statusSummary.Alpha / total) * 100).toFixed(1)
                  : "0";
              const pctIzin =
                total > 0
                  ? ((statusSummary.Izin / total) * 100).toFixed(1)
                  : "0";
              const pctSakit =
                total > 0
                  ? ((statusSummary.Sakit / total) * 100).toFixed(1)
                  : "0";
              return (
                <div className="mb-4 p-3 bg-gray-50 border border-gray-200 rounded-lg">
                  <p className="text-center text-sm font-semibold text-gray-700 mb-2">
                    Persentase Kehadiran {selectedBulan} — Kelas {selectedKelas}
                  </p>
                  <div className="grid grid-cols-4 gap-2 text-center text-sm">
                    <div>
                      <div className="text-green-600 font-bold text-base">
                        {pctHadir}%
                      </div>
                      <div className="text-gray-500">Hadir</div>
                    </div>
                    <div>
                      <div className="text-yellow-600 font-bold text-base">
                        {pctIzin}%
                      </div>
                      <div className="text-gray-500">Izin</div>
                    </div>
                    <div>
                      <div className="text-blue-600 font-bold text-base">
                        {pctSakit}%
                      </div>
                      <div className="text-gray-500">Sakit</div>
                    </div>
                    <div>
                      <div className="text-red-600 font-bold text-base">
                        {pctAlpha}%
                      </div>
                      <div className="text-gray-500">Alpha</div>
                    </div>
                  </div>
                </div>
              );
            })()}
            <div className="overflow-x-auto">
              <table className="min-w-full border-collapse border border-gray-200">
                <thead>
                  <tr className="bg-gray-100">
                    <th className="border border-gray-200 px-2 py-0.5 text-center text-sm font-semibold text-gray-700">
                      No.
                    </th>
                    <th className="border border-gray-200 px-1 py-0.5 text-left text-sm font-semibold text-gray-700">
                      Nama
                    </th>
                    <th className="border border-gray-200 px-1 py-0.5 text-left text-sm font-semibold text-gray-700">
                      Kelas
                    </th>
                    <th className="border border-gray-200 px-1 py-0.5 text-center text-sm font-semibold text-gray-700">
                      Hadir
                    </th>
                    <th className="border border-gray-200 px-1 py-0.5 text-center text-sm font-semibold text-gray-700">
                      Alpha
                    </th>
                    <th className="border border-gray-200 px-1 py-0.5 text-center text-sm font-semibold text-gray-700">
                      Izin
                    </th>
                    <th className="border border-gray-200 px-1 py-0.5 text-center text-sm font-semibold text-gray-700">
                      Sakit
                    </th>
                    <th className="border border-gray-200 px-1 py-0.5 text-center text-sm font-semibold text-gray-700">
                      % Hadir
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {filteredRecapData.map((item, index) => (
                    <tr
                      key={index}
                      className={index % 2 === 0 ? "bg-white" : "bg-gray-50"}
                    >
                      <td className="border border-gray-200 px-2 py-0.5 text-center text-sm text-gray-600 font-medium">
                        {index + 1}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-sm text-gray-600">
                        {item.nama || "N/A"}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-sm text-gray-600">
                        {item.kelas || "N/A"}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-center text-sm text-gray-600">
                        {item.hadir || 0}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-center text-sm text-gray-600">
                        {item.alpa || 0}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-center text-sm text-gray-600">
                        {item.izin || 0}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-center text-sm text-gray-600">
                        {item.sakit || 0}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-center text-sm text-gray-600">
                        {item.persenHadir !== undefined
                          ? `${item.persenHadir}%`
                          : "N/A"}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div className="mt-6 flex gap-4 justify-center">
              <button
                onClick={downloadExcel}
                className="px-1 py-0.5 bg-green-600 hover:bg-green-700 text-white rounded-lg font-medium transition-colors"
              >
                📥 Download Excel
              </button>
              <button
                onClick={downloadPDF}
                className="px-1 py-0.5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg font-medium transition-colors"
              >
                📄 Download PDF
              </button>
            </div>
          </>
        )}
      </div>
    </div>
  );
};

const GraphTab: React.FC<{
  uniqueClasses: string[];
}> = ({ uniqueClasses }) => {
  const [graphData, setGraphData] = useState<GraphData>({
    Januari: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
    Februari: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
    Maret: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
    April: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
    Mei: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
    Juni: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
    Juli: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
    Agustus: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
    September: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
    Oktober: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
    November: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
    Desember: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
  });
  const [selectedKelas, setSelectedKelas] = useState<string>(
    uniqueClasses.find((k) => k !== "Semua") || ""
  );
  const [selectedSemester, setSelectedSemester] = useState<"1" | "2">("2");
  const [statusVisibility, setStatusVisibility] = useState<StatusVisibility>({
    Hadir: true,
    Alpha: true,
    Izin: true,
    Sakit: true,
  });
  const [loading, setLoading] = useState<boolean>(true);
  const [semesterSummary, setSemesterSummary] = useState({
    Hadir: 0,
    Alpha: 0,
    Izin: 0,
    Sakit: 0,
  });

  const uniqueClassesWithDefault = React.useMemo(() => {
    return uniqueClasses.filter((kelas) => kelas !== "Semua");
  }, [uniqueClasses]);

  useEffect(() => {
    setLoading(true);

    // Fetch 1: untuk jumlah & persentase (dari semesterRecap)
    fetch(
      `${endpoint}?action=semesterRecap&kelas=${selectedKelas}&semester=${selectedSemester}`
    )
      .then((res) => {
        if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
        return res.json();
      })
      .then((data) => {
        if (data.success) {
          const recap = data.data || [];
          const totals = recap.reduce(
            (
              acc: {
                Hadir: number;
                Alpha: number;
                Izin: number;
                Sakit: number;
              },
              item: any
            ) => {
              acc.Hadir += item.hadir || 0;
              acc.Alpha += item.alpa || 0;
              acc.Izin += item.izin || 0;
              acc.Sakit += item.sakit || 0;
              return acc;
            },
            { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 }
          );
          setSemesterSummary(totals);
        }
      })
      .catch((error) => console.error("Error fetch semesterRecap:", error));

    // Fetch 2: untuk grafik batang (dari graphData)
    fetch(
      `${endpoint}?action=graphData&kelas=${selectedKelas}&semester=${selectedSemester}`
    )
      .then((res) => {
        if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
        return res.json();
      })
      .then((data) => {
        if (data.success) {
          setGraphData(data.data || {});
        } else {
          alert("❌ Gagal memuat data grafik: " + data.message);
        }
        setLoading(false);
      })
      .catch((error) => {
        console.error("Error fetch graphData:", error);
        alert("❌ Gagal memuat data grafik. Cek console untuk detail.");
        setLoading(false);
      });
  }, [selectedKelas, selectedSemester]);

  const semesterMonths: Record<"1" | "2", string[]> = {
    "1": ["Juli", "Agustus", "September", "Oktober", "November", "Desember"],
    "2": ["Januari", "Februari", "Maret", "April", "Mei", "Juni"],
  };

  const chartData: ChartData<"bar", number[], string> = {
    labels: semesterMonths[selectedSemester],
    datasets: [
      ...(statusVisibility.Hadir
        ? [
            {
              label: "Hadir",
              data: semesterMonths[selectedSemester].map(
                (month: string) => graphData[month]?.Hadir || 0
              ),
              backgroundColor: "rgba(75, 192, 192, 0.6)",
              borderColor: "rgba(75, 192, 192, 1)",
              borderWidth: 1,
            },
          ]
        : []),
      ...(statusVisibility.Alpha
        ? [
            {
              label: "Alpha",
              data: semesterMonths[selectedSemester].map(
                (month: string) => graphData[month]?.Alpha || 0
              ),
              backgroundColor: "rgba(255, 99, 132, 0.6)",
              borderColor: "rgba(255, 99, 132, 1)",
              borderWidth: 1,
            },
          ]
        : []),
      ...(statusVisibility.Izin
        ? [
            {
              label: "Izin",
              data: semesterMonths[selectedSemester].map(
                (month: string) => graphData[month]?.Izin || 0
              ),
              backgroundColor: "rgba(255, 205, 86, 0.6)",
              borderColor: "rgba(255, 205, 86, 1)",
              borderWidth: 1,
            },
          ]
        : []),
      ...(statusVisibility.Sakit
        ? [
            {
              label: "Sakit",
              data: semesterMonths[selectedSemester].map(
                (month: string) => graphData[month]?.Sakit || 0
              ),
              backgroundColor: "rgba(54, 162, 235, 0.6)",
              borderColor: "rgba(54, 162, 235, 1)",
              borderWidth: 1,
            },
          ]
        : []),
    ],
  };

  const chartOptions: ChartOptions<"bar"> = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: {
        position: "top" as const,
        onClick: (
          e: ChartEvent,
          legendItem: LegendItem,
          legend: {
            chart: {
              data: { datasets: { hidden?: boolean }[] };
              update: () => void;
            };
          }
        ) => {
          const index = legendItem.datasetIndex;
          if (index !== undefined) {
            const ci = legend.chart.data.datasets[index];
            ci.hidden = !ci.hidden;
            legend.chart.update();
            setStatusVisibility((prev) => ({
              ...prev,
              [legendItem.text as keyof StatusVisibility]: !ci.hidden,
            }));
          }
        },
      },
      title: {
        display: true,
        text: `Persentase Kehadiran Kelas ${selectedKelas} Semester ${selectedSemester} 2025`,
      },
    },
    scales: {
      y: {
        beginAtZero: true,
        max: 100,
        ticks: {
          stepSize: 10,
          font: {
            size: 10,
          },
          autoSkip: false,
          maxTicksLimit: 11,
        },
        title: { display: true, text: "Persentase (%)" },
      },
      x: {
        ticks: {
          font: {
            size: 10,
          },
        },
      },
    },
  };

  const handleStatusToggle = (status: keyof StatusVisibility) => {
    setStatusVisibility((prev) => ({
      ...prev,
      [status]: !prev[status],
    }));
  };

  return (
    <div className="max-w-4xl mx-auto" style={{ paddingBottom: "70px" }}>
      <div className="bg-white p-6 rounded-lg shadow-md">
        <h2 className="text-2xl font-bold text-center text-blue-700 mb-6">
          📈 Grafik Kehadiran
        </h2>

        <div className="mb-6 flex flex-col md:flex-row gap-4 items-center justify-center">
          <div className="text-center">
            <p className="text-sm text-gray-500 mb-2">Filter Kelas</p>
            <select
              value={selectedKelas}
              onChange={(e) => setSelectedKelas(e.target.value)}
              className="border border-gray-300 rounded-lg px-1 py-0.5 shadow-sm bg-white min-w-32"
            >
              {uniqueClassesWithDefault.map((kelas) => (
                <option key={kelas} value={kelas}>
                  {kelas}
                </option>
              ))}
            </select>
          </div>
          <div className="text-center">
            <p className="text-sm text-gray-500 mb-2">Filter Semester</p>
            <select
              value={selectedSemester}
              onChange={(e) => setSelectedSemester(e.target.value as "1" | "2")}
              className="border border-gray-300 rounded-lg px-1 py-0.5 shadow-sm bg-white min-w-32"
            >
              <option value="1">Semester 1 (Juli-Des)</option>
              <option value="2">Semester 2 (Jan-Jun)</option>
            </select>
          </div>
        </div>

        <div className="mb-6 flex flex-wrap gap-4 justify-center">
          {(["Hadir", "Alpha", "Izin", "Sakit"] as const).map((status) => (
            <label key={status} className="flex items-center gap-2">
              <input
                type="checkbox"
                checked={statusVisibility[status]}
                onChange={() => handleStatusToggle(status)}
                className="h-4 w-4 text-blue-600 rounded"
              />
              <span className="text-sm text-gray-700">{status}</span>
            </label>
          ))}
        </div>

        {loading ? (
          <div className="text-center py-8">
            <p className="text-gray-500">Memuat grafik...</p>
          </div>
        ) : (
          <>
            {(() => {
              const total =
                semesterSummary.Hadir +
                semesterSummary.Alpha +
                semesterSummary.Izin +
                semesterSummary.Sakit;
              const pctHadir =
                total > 0
                  ? ((semesterSummary.Hadir / total) * 100).toFixed(1)
                  : "0";
              const pctAlpha =
                total > 0
                  ? ((semesterSummary.Alpha / total) * 100).toFixed(1)
                  : "0";
              const pctIzin =
                total > 0
                  ? ((semesterSummary.Izin / total) * 100).toFixed(1)
                  : "0";
              const pctSakit =
                total > 0
                  ? ((semesterSummary.Sakit / total) * 100).toFixed(1)
                  : "0";
              return (
                <div className="mb-4 p-3 bg-gray-50 border border-gray-200 rounded-lg">
                  <p className="text-center text-sm font-semibold text-gray-700 mb-3">
                    Rekap Kehadiran Kelas {selectedKelas} — Semester{" "}
                    {selectedSemester}
                  </p>
                  <div className="grid grid-cols-4 gap-2 text-center text-sm mb-2">
                    <div className="bg-green-50 border border-green-200 rounded-lg p-2">
                      <div className="text-green-600 font-bold text-base">
                        {semesterSummary.Hadir}
                      </div>
                      <div className="text-gray-500 text-xs">Hadir</div>
                    </div>
                    <div className="bg-yellow-50 border border-yellow-200 rounded-lg p-2">
                      <div className="text-yellow-600 font-bold text-base">
                        {semesterSummary.Izin}
                      </div>
                      <div className="text-gray-500 text-xs">Izin</div>
                    </div>
                    <div className="bg-blue-50 border border-blue-200 rounded-lg p-2">
                      <div className="text-blue-600 font-bold text-base">
                        {semesterSummary.Sakit}
                      </div>
                      <div className="text-gray-500 text-xs">Sakit</div>
                    </div>
                    <div className="bg-red-50 border border-red-200 rounded-lg p-2">
                      <div className="text-red-600 font-bold text-base">
                        {semesterSummary.Alpha}
                      </div>
                      <div className="text-gray-500 text-xs">Alpha</div>
                    </div>
                  </div>
                  <div className="grid grid-cols-4 gap-2 text-center text-sm">
                    <div>
                      <div className="text-green-600 font-bold">
                        {pctHadir}%
                      </div>
                      <div className="text-gray-400 text-xs">Hadir</div>
                    </div>
                    <div>
                      <div className="text-yellow-600 font-bold">
                        {pctIzin}%
                      </div>
                      <div className="text-gray-400 text-xs">Izin</div>
                    </div>
                    <div>
                      <div className="text-blue-600 font-bold">{pctSakit}%</div>
                      <div className="text-gray-400 text-xs">Sakit</div>
                    </div>
                    <div>
                      <div className="text-red-600 font-bold">{pctAlpha}%</div>
                      <div className="text-gray-400 text-xs">Alpha</div>
                    </div>
                  </div>
                </div>
              );
            })()}
            <div
              className="h-96"
              style={{
                minHeight: "300px",
                maxHeight: "500px",
              }}
            >
              <Bar data={chartData} options={chartOptions} />
            </div>
          </>
        )}
      </div>
    </div>
  );
};

const SemesterRecapTab: React.FC<{
  uniqueClasses: string[];
  students: Student[]; // ✅ TAMBAHKAN INI
}> = ({
  uniqueClasses,
  students, // ✅ DAN INI
}) => {
  const [recapData, setRecapData] = useState<SemesterRecap[]>([]);
  const [selectedKelas, setSelectedKelas] = useState<string>(
    uniqueClasses.find((k) => k !== "Semua") || ""
  );
  const [selectedSemester, setSelectedSemester] = useState<"1" | "2">("1");
  const [selectedDate, setSelectedDate] = useState<string>(
    new Date().toISOString().split("T")[0]
  );

  const [loading, setLoading] = useState<boolean>(true);
  const [showPreviewModal, setShowPreviewModal] = useState<boolean>(false);
  const [schoolData, setSchoolData] = useState<SchoolData | null>(null);
  const [downloadingPreviewExcel, setDownloadingPreviewExcel] =
    useState<boolean>(false);

  useEffect(() => {
    setLoading(true);
    const sheetName =
      selectedSemester === "1" ? SHEET_SEMESTER1 : SHEET_SEMESTER2;
    fetch(
      `${endpoint}?action=semesterRecap&kelas=${
        selectedKelas === "Semua" ? "" : selectedKelas
      }&semester=${selectedSemester}`
    )
      .then((res) => {
        if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
        return res.json();
      })
      .then((data) => {
        if (data.success) {
          setRecapData(data.data || []);
        } else {
          alert(
            `❌ Gagal memuat data rekap ${
              selectedSemester === "1" ? "Semester 1" : "Semester 2"
            }: ${data.message}`
          );
          setRecapData([]);
        }
        setLoading(false);
      })
      .catch((error) => {
        console.error("Error fetch:", error);
        alert(
          `❌ Gagal memuat data rekap ${
            selectedSemester === "1" ? "Semester 1" : "Semester 2"
          }. Cek console untuk detail.`
        );
        setRecapData([]);
        setLoading(false);
      });

    fetch(`${endpoint}?action=schoolData`)
      .then((res) => {
        if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
        return res.json();
      })
      .then((data) => {
        if (data.success && data.data && data.data.length > 0) {
          setSchoolData(data.data[0]);
        } else {
          setSchoolData(null);
        }
      })
      .catch((error) => {
        console.error("Error fetching school data:", error);
        alert("❌ Gagal memuat data sekolah. Cek console untuk detail.");
      });
  }, [selectedKelas, selectedSemester]);

  const filteredRecapData = React.useMemo(() => {
    if (!selectedKelas) return recapData;
    return recapData.filter(
      (item) => String(item.kelas).trim() === selectedKelas
    );
  }, [recapData, selectedKelas]);

  const getStatusSummary = (): {
    Hadir: number;
    Izin: number;
    Sakit: number;
    Alpha: number;
  } => {
    const summary = { Hadir: 0, Izin: 0, Sakit: 0, Alpha: 0 };
    filteredRecapData.forEach((item) => {
      summary.Hadir += item.hadir || 0;
      summary.Alpha += item.alpa || 0;
      summary.Izin += item.izin || 0;
      summary.Sakit += item.sakit || 0;
    });
    return summary;
  };

  const statusSummary = getStatusSummary();
  const getPotongan = (jumlah: number): number => {
    if (jumlah <= 3) return 0;
    if (jumlah <= 6) return 3;
    if (jumlah <= 15) return 6;
    if (jumlah <= 20) return 10;
    return 15;
  };

  const downloadPreviewExcel = async () => {
    setDownloadingPreviewExcel(true);

    try {
      const response = await fetch(
        `${endpoint}?action=generatePreviewExcel&kelas=${encodeURIComponent(
          selectedKelas
        )}&semester=${selectedSemester}`
      );
      const result = await response.json();

      if (result.success) {
        window.open(result.url, "_blank");
      } else {
        alert("❌ Gagal generate Excel: " + (result.message || result.error));
      }
    } catch (err) {
      alert(
        "❌ Error: " + (err instanceof Error ? err.message : "Unknown error")
      );
    } finally {
      setDownloadingPreviewExcel(false);
    }
  };

  const downloadPreviewPDF = () => {
    const doc = new jsPDF();
    const pageWidth = doc.internal.pageSize.getWidth();
    const margin = 14;
    let currentY = margin;

    doc.setFont("Times", "roman");

    const namaSekolah = schoolData?.namaSekolah || "UPT SDN 13 BATANG";
    const semesterLabel =
      selectedSemester === "1" ? "SEMESTER 1" : "SEMESTER 2";
    const title = `PREVIEW PENGURANGAN KEHADIRAN KELAS ${selectedKelas}  ${namaSekolah}  ${semesterLabel}`;

    doc.setFontSize(11);
    doc.setFont("Times", "bold");
    doc.text(title, pageWidth / 2, currentY, { align: "center" });
    currentY += 8;

    const headers = [
      "No.",
      "Nama",
      "Alpha",
      "Pot. Alpha",
      "Izin",
      "Pot. Izin",
      "Sakit",
      "Pot. Sakit",
    ];

    const body = filteredRecapData.map((item, index) => {
      const alpha = item.alpa || 0;
      const izin = item.izin || 0;
      const sakit = item.sakit || 0;
      return [
        index + 1,
        item.nama || "N/A",
        alpha,
        getPotongan(alpha),
        izin,
        getPotongan(izin),
        sakit,
        getPotongan(sakit),
      ];
    });

    autoTable(doc, {
      head: [headers],
      body,
      startY: currentY,
      styles: { font: "Times", fontSize: 8, cellPadding: 2, halign: "center" },
      headStyles: {
        fillColor: [255, 255, 0],
        textColor: [0, 0, 0],
        fontStyle: "bold",
      },
      columnStyles: {
        1: { halign: "left", cellWidth: 45 },
      },
      alternateRowStyles: { fillColor: [240, 240, 240] },
      didParseCell: (data) => {
        // Highlight kolom pengurangan (index 3, 5, 7)
        if ([3, 5, 7].includes(data.column.index) && data.section === "body") {
          data.cell.styles.fillColor = [255, 251, 235];
          data.cell.styles.textColor = [161, 98, 7];
          data.cell.styles.fontStyle = "bold";
        }
      },
    });

    const date = new Date()
      .toLocaleString("id-ID", {
        day: "2-digit",
        month: "long",
        year: "numeric",
        hour: "2-digit",
        minute: "2-digit",
        hour12: false,
      })
      .replace(/ /g, "_")
      .replace(/:/g, "-");

    const fileName = `Preview_Pengurangan_${selectedSemester}_${selectedKelas}_${date}.pdf`;
    doc.save(fileName);
  };

  const downloadExcel = () => {
    const headers = [
      "No.",
      "Nama",
      "Kelas",
      "Hadir",
      "Alpha",
      "Izin",
      "Sakit",
      "% Hadir",
    ];
    const data = [
      headers,
      ...filteredRecapData.map((item, index) => [
        index + 1, // Nomor urut
        item.nama || "N/A",
        item.kelas || "N/A",
        item.hadir || 0,
        item.alpa || 0,
        item.izin || 0,
        item.sakit || 0,
        item.persenHadir !== undefined ? `${item.persenHadir}%` : "N/A",
      ]),
      [
        "",
        "TOTAL",
        "",
        statusSummary.Hadir,
        statusSummary.Alpha,
        statusSummary.Izin,
        statusSummary.Sakit,
        "",
      ],
      [
        "",
        "PERSEN",
        "",
        `${(
          (statusSummary.Hadir /
            (statusSummary.Hadir +
              statusSummary.Alpha +
              statusSummary.Izin +
              statusSummary.Sakit)) *
          100
        ).toFixed(2)}%`,
        `${(
          (statusSummary.Alpha /
            (statusSummary.Hadir +
              statusSummary.Alpha +
              statusSummary.Izin +
              statusSummary.Sakit)) *
          100
        ).toFixed(2)}%`,
        `${(
          (statusSummary.Izin /
            (statusSummary.Hadir +
              statusSummary.Alpha +
              statusSummary.Izin +
              statusSummary.Sakit)) *
          100
        ).toFixed(2)}%`,
        `${(
          (statusSummary.Sakit /
            (statusSummary.Hadir +
              statusSummary.Alpha +
              statusSummary.Izin +
              statusSummary.Sakit)) *
          100
        ).toFixed(2)}%`,
        "",
      ],
    ];

    const ws = XLSX.utils.aoa_to_sheet(data);
    ws["!cols"] = [
      { wch: 5 }, // Lebar kolom No. (sempit)
      { wch: 25 }, // Nama
      { wch: 10 }, // Kelas
      { wch: 10 }, // Hadir
      { wch: 10 }, // Alpha
      { wch: 10 }, // Izin
      { wch: 10 }, // Sakit
      { wch: 10 }, // % Hadir
    ];
    const headerStyle = {
      font: { bold: true },
      fill: { fgColor: { rgb: "FFFF00" } },
      alignment: { horizontal: "center" },
    };
    const totalStyle = {
      font: { bold: true },
      fill: { fgColor: { rgb: "D3D3D3" } },
      alignment: { horizontal: "center" },
    };
    const percentStyle = {
      font: { bold: true },
      fill: { fgColor: { rgb: "D3D3D3" } },
      alignment: { horizontal: "center" },
    };
    headers.forEach((header, index) => {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: index });
      ws[cellAddress] = { ...ws[cellAddress], s: headerStyle };
    });
    const totalRow = filteredRecapData.length + 1;
    ["A", "B", "C", "D", "E", "F", "G", "H"].forEach((col, idx) => {
      const cellAddress = `${col}${totalRow}`;
      ws[cellAddress] = { ...ws[cellAddress], s: totalStyle };
    });
    const percentRow = filteredRecapData.length + 2;
    ["A", "B", "C", "D", "E", "F", "G", "H"].forEach((col, idx) => {
      const cellAddress = `${col}${percentRow}`;
      ws[cellAddress] = { ...ws[cellAddress], s: percentStyle };
    });

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Rekap Semester");
    const date = new Date()
      .toLocaleString("id-ID", {
        day: "2-digit",
        month: "long",
        year: "numeric",
        hour: "2-digit",
        minute: "2-digit",
        hour12: false,
      })
      .replace(/ /g, "_")
      .replace(/:/g, "-");
    const fileName = `Rekap_Semester_${selectedSemester}_${selectedKelas}_${date}.xlsx`;
    XLSX.writeFile(wb, fileName);
  };

  const downloadPDF = async () => {
    const doc = new jsPDF();
    const pageWidth = doc.internal.pageSize.getWidth();
    const margin = 14;
    const lineSpacing = 5;
    let currentY = margin;

    doc.setFont("Times", "roman");

    // Title - Format sama dengan Daftar Hadir
    const namaSekolah = schoolData?.namaSekolah || "UPT SDN 13 BATANG";

    // ✅ TAMBAHKAN: Ambil tahun dari selectedDate
    const tahunDariTanggal = new Date(selectedDate).getFullYear();

    // Judul dalam 1 baris dengan format Semester - gunakan tahunDariTanggal
    const semesterLabel =
      selectedSemester === "1" ? "SEMESTER 1" : "SEMESTER 2";
    const title = `REKAP ABSENSI SISWA KELAS ${selectedKelas}  ${namaSekolah}  ${semesterLabel} ${tahunDariTanggal}`;

    doc.setFontSize(12); // Ukuran font sama dengan daftar hadir
    doc.setFont("Times", "bold");
    doc.text(title, pageWidth / 2, currentY, { align: "center" });

    currentY += 10;

    const headers = [
      "No.",
      "Nama",
      "Kelas",
      "Hadir",
      "Alpha",
      "Izin",
      "Sakit",
      "% Hadir",
    ];
    const body = filteredRecapData.map((item, index) => [
      index + 1, // Nomor urut
      item.nama || "N/A",
      item.kelas || "N/A",
      item.hadir || 0,
      item.alpa || 0,
      item.izin || 0,
      item.sakit || 0,
      item.persenHadir !== undefined ? `${item.persenHadir}%` : "N/A",
    ]);

    const totalRow = [
      "",
      "TOTAL",
      "",
      statusSummary.Hadir,
      statusSummary.Alpha,
      statusSummary.Izin,
      statusSummary.Sakit,
      "",
    ];

    const percentRow = [
      "",
      "PERSEN",
      "",
      `${(
        (statusSummary.Hadir /
          (statusSummary.Hadir +
            statusSummary.Alpha +
            statusSummary.Izin +
            statusSummary.Sakit)) *
        100
      ).toFixed(2)}%`,
      `${(
        (statusSummary.Alpha /
          (statusSummary.Hadir +
            statusSummary.Alpha +
            statusSummary.Izin +
            statusSummary.Sakit)) *
        100
      ).toFixed(2)}%`,
      `${(
        (statusSummary.Izin /
          (statusSummary.Hadir +
            statusSummary.Alpha +
            statusSummary.Izin +
            statusSummary.Sakit)) *
        100
      ).toFixed(2)}%`,
      `${(
        (statusSummary.Sakit /
          (statusSummary.Hadir +
            statusSummary.Alpha +
            statusSummary.Izin +
            statusSummary.Sakit)) *
        100
      ).toFixed(2)}%`,
      "",
    ];

    autoTable(doc, {
      head: [headers],
      body: [...body, totalRow, percentRow],
      startY: currentY,
      styles: { font: "Times", fontSize: 8, cellPadding: 2 },
      headStyles: {
        fillColor: [255, 255, 0],
        textColor: [0, 0, 0],
        fontStyle: "bold",
      },
      alternateRowStyles: { fillColor: [240, 240, 240] },
      columnStyles: {
        0: { cellWidth: 10 }, // No. (sempit)
        1: { cellWidth: 50 }, // Nama
        2: { cellWidth: 20 }, // Kelas
        3: { cellWidth: 20 }, // Hadir
        4: { cellWidth: 20 }, // Alpha
        5: { cellWidth: 20 }, // Izin
        6: { cellWidth: 20 }, // Sakit
        7: { cellWidth: 20 }, // % Hadir
      },
    });

    currentY = (doc as any).lastAutoTable.finalY + 10;

    // ✅ TAMBAHAN BARU: Tabel Informasi Jumlah Siswa
    const pageHeight = doc.internal.pageSize.getHeight();
    const bottomMargin = 20;
    const spaceNeededForStudentTable = 20;
    const spaceNeededForSignatures = 60; // Ruang untuk tanda tangan

    // ✅ CEK APAKAH TANDA TANGAN + TABEL SISWA MUAT DI HALAMAN INI
    // Cek apakah ada cukup ruang untuk tabel jumlah siswa DAN tanda tangan
    if (
      currentY + spaceNeededForStudentTable + spaceNeededForSignatures >
      pageHeight - bottomMargin
    ) {
      doc.addPage();
      currentY = margin;
    }

    // Hitung jumlah siswa berdasarkan jenis kelamin dari filteredRecapData
    const genderSummary = filteredRecapData.reduce(
      (acc, student) => {
        // Cari data lengkap siswa untuk mendapatkan jenis kelamin
        const studentData = students.find((s) => s.name === student.nama);
        if (studentData) {
          const jenisKelamin = String(studentData.jenisKelamin || "")
            .trim()
            .toUpperCase();
          if (jenisKelamin === "L" || jenisKelamin === "LAKI-LAKI") {
            acc.lakiLaki++;
          } else if (jenisKelamin === "P" || jenisKelamin === "PEREMPUAN") {
            acc.perempuan++;
          }
        }
        return acc;
      },
      { lakiLaki: 0, perempuan: 0 }
    );

    const totalSiswa = genderSummary.lakiLaki + genderSummary.perempuan;

    doc.setFontSize(10);
    doc.setFont("Times", "bold");
    doc.text("JUMLAH SISWA", margin, currentY, { align: "left" });
    currentY += 3;

    const tableWidth = (pageWidth - 2 * margin) * 0.4;

    autoTable(doc, {
      head: [["LAKI-LAKI", "PEREMPUAN", "TOTAL SISWA"]],
      body: [
        [
          genderSummary.lakiLaki.toString(),
          genderSummary.perempuan.toString(),
          totalSiswa.toString(),
        ],
      ],
      startY: currentY,
      margin: { left: margin, right: pageWidth - margin - tableWidth },
      tableWidth: tableWidth,
      theme: "grid",
      styles: {
        font: "Times",
        fontSize: 7,
        cellPadding: 1,
        halign: "center",
        valign: "middle",
        lineWidth: 0.5,
      },
      headStyles: {
        fillColor: [255, 255, 255],
        textColor: [0, 0, 0],
        fontStyle: "bold",
        lineWidth: 1,
      },
      bodyStyles: {
        fillColor: [255, 255, 255],
        textColor: [0, 0, 0],
        fontStyle: "bold",
        fontSize: 10,
        lineWidth: 1,
      },
      columnStyles: {
        0: { cellWidth: tableWidth / 3, fillColor: [255, 255, 255] },
        1: { cellWidth: tableWidth / 3, fillColor: [255, 255, 255] },
        2: { cellWidth: tableWidth / 3, fillColor: [255, 255, 255] },
      },
    });

    currentY = (doc as any).lastAutoTable.finalY + 10;

    if (schoolData) {
      doc.setFontSize(10);
      doc.setFont("Times", "roman");

      const formattedDate = new Date(selectedDate).toLocaleDateString("id-ID", {
        day: "2-digit",
        month: "long",
        year: "numeric",
      });
      const placeDateText = `${
        schoolData.namaKota || "Makassar"
      }, ${formattedDate}`;
      const rightColumnX = pageWidth - margin - 50; // Signature width is 50
      doc.text(placeDateText, rightColumnX + 25, currentY - 1, {
        align: "center",
      });
      currentY += 5; // Keep close to "Guru Kelas"

      const principalText = [
        "Kepala Sekolah,",
        "",
        "",
        `( ${schoolData.namaKepsek || "N/A"} )`,
        `NIP: ${schoolData.nipKepsek || "N/A"}`,
      ];
      const teacherText = [
        "Guru Kelas,",
        "",
        "",
        `( ${schoolData.namaGuru || "N/A"} )`,
        `NIP: ${schoolData.nipGuru || "N/A"}`,
      ];

      const signatureWidth = 30;
      const signatureHeight = 20;
      const leftColumnX = margin;

      // Principal signature and text
      if (schoolData.ttdKepsek) {
        try {
          const canvas = document.createElement("canvas");
          canvas.width = 150; // Sesuaikan ukuran canvas (lebar lebih besar untuk tanda tangan panjang)
          canvas.height = 50; // Sesuaikan ukuran canvas (tinggi cukup untuk garis tanda tangan)
          const ctx = canvas.getContext("2d");
          const v = await Canvg.from(ctx, schoolData.ttdKepsek); // schoolData.ttdKepsek adalah base64 SVG
          v.start();
          const pngData = canvas.toDataURL("image/png");
          doc.addImage(
            pngData,
            "PNG",
            leftColumnX + 10,
            currentY - 3,
            signatureWidth,
            signatureHeight
          ); // Sesuaikan posisi sesuai asli
        } catch (error) {
          console.error("Error rendering Kepsek signature:", error);
          doc.setFontSize(10);
          doc.text(
            "Gagal render tanda tangan Kepala Sekolah.",
            leftColumnX + 10,
            currentY - 3 + 10
          );
        }
      }

      // Pisahkan "Kepala Sekolah" dengan posisi yang lebih tinggi
      doc.text("Kepala Sekolah,", leftColumnX + 25, currentY - 2, {
        align: "center",
      });

      // Kosong dan kosong
      doc.text("", leftColumnX + 25, currentY + lineSpacing, {
        align: "center",
      });
      doc.text("", leftColumnX + 25, currentY + 2 * lineSpacing, {
        align: "center",
      });

      // Nama kepala sekolah dengan format bold dan underline
      const principalName = schoolData.namaKepsek || "N/A";
      doc.setFont("Times", "bold");
      doc.text(principalName, leftColumnX + 25, currentY + 3.5 * lineSpacing, {
        align: "center",
      });

      // Add underline to principal name
      const principalNameText = principalName;
      const textWidth = doc.getTextWidth(principalNameText);
      const textX = leftColumnX + 25 - textWidth / 2;
      doc.line(
        textX,
        currentY + 3.5 * lineSpacing + 1,
        textX + textWidth,
        currentY + 3.5 * lineSpacing + 1
      );

      // Reset font and add NIP
      doc.setFont("Times", "roman");
      doc.text(
        `NIP. ${schoolData.nipKepsek || "N/A"}`,
        leftColumnX + 25,
        currentY + 4.5 * lineSpacing,
        {
          align: "center",
        }
      );

      // Teacher signature and text
      if (schoolData.ttdGuru) {
        try {
          const canvas = document.createElement("canvas");
          canvas.width = 150; // Sesuaikan ukuran canvas
          canvas.height = 50;
          const ctx = canvas.getContext("2d");
          const v = await Canvg.from(ctx, schoolData.ttdGuru); // schoolData.ttdGuru adalah base64 SVG
          v.start();
          const pngData = canvas.toDataURL("image/png");
          doc.addImage(
            pngData,
            "PNG",
            rightColumnX + 10,
            currentY - 5,
            signatureWidth,
            signatureHeight
          ); // Sesuaikan posisi sesuai asli
        } catch (error) {
          console.error("Error rendering Guru signature:", error);
          doc.setFontSize(10);
          doc.text(
            "Gagal render tanda tangan Guru.",
            rightColumnX + 10,
            currentY - 5 + 10
          );
        }
      }

      // Pisahkan "Guru Kelas" dengan posisi yang lebih tinggi
      doc.text(
        `${schoolData.statusGuru || "Guru Kelas"},`,
        rightColumnX + 25,
        currentY - 2,
        {
          align: "center",
        }
      );

      // Kosong dan kosong
      doc.text("", rightColumnX + 25, currentY + lineSpacing, {
        align: "center",
      });
      doc.text("", rightColumnX + 25, currentY + 2 * lineSpacing, {
        align: "center",
      });

      // Nama guru dengan format bold dan underline
      const teacherName = schoolData.namaGuru || "N/A";
      doc.setFont("Times", "bold");
      doc.text(teacherName, rightColumnX + 25, currentY + 3.5 * lineSpacing, {
        align: "center",
      });

      // Add underline to teacher name
      const teacherNameText = teacherName;
      const teacherTextWidth = doc.getTextWidth(teacherNameText);
      const teacherTextX = rightColumnX + 25 - teacherTextWidth / 2;
      doc.line(
        teacherTextX,
        currentY + 3.5 * lineSpacing + 1,
        teacherTextX + teacherTextWidth,
        currentY + 3.5 * lineSpacing + 1
      );

      // Reset font and add NIP
      doc.setFont("Times", "roman");
      doc.text(
        `NIP. ${schoolData.nipGuru || "N/A"}`,
        rightColumnX + 25,
        currentY + 4.5 * lineSpacing,
        {
          align: "center",
        }
      );
    } else {
      doc.setFontSize(10);
      doc.text("Data sekolah tidak tersedia.", margin, currentY);
    }

    const date = new Date()
      .toLocaleString("id-ID", {
        day: "2-digit",
        month: "long",
        year: "numeric",
        hour: "2-digit",
        minute: "2-digit",
        hour12: false,
      })
      .replace(/ /g, "_")
      .replace(/:/g, "-");
    const fileName = `Rekap_Semester_${selectedSemester}_${selectedKelas}_${date}.pdf`;
    doc.save(fileName);
  };

  return (
    <div className="max-w-4xl mx-auto" style={{ paddingBottom: "70px" }}>
      <div className="bg-white p-6 rounded-lg shadow-md">
        <h2 className="text-2xl font-bold text-center text-blue-700 mb-6">
          📊 Rekap Absensi Semester
        </h2>

        <div className="mb-6 flex flex-col md:flex-row gap-4 items-center justify-center">
          <div className="text-center">
            <p className="text-sm text-gray-500 mb-2">Filter Kelas</p>
            <select
              value={selectedKelas}
              onChange={(e) => setSelectedKelas(e.target.value)}
              className="border border-gray-300 rounded-lg px-1 py-0.5 shadow-sm bg-white min-w-32"
            >
              {uniqueClasses
                .filter((kelas) => kelas !== "Semua")
                .map((kelas) => (
                  <option key={kelas} value={kelas}>
                    {kelas}
                  </option>
                ))}
            </select>
          </div>
          <div className="text-center">
            <p className="text-sm text-gray-500 mb-2">Pilih Semester</p>
            <select
              value={selectedSemester}
              onChange={(e) => setSelectedSemester(e.target.value as "1" | "2")}
              className="border border-gray-300 rounded-lg px-1 py-0.5 shadow-sm bg-white min-w-32"
            >
              <option value="1">Semester 1</option>
              <option value="2">Semester 2</option>
            </select>
          </div>
        </div>

        {/* Separator line and PDF settings section */}
        <div className="border-t border-gray-200 pt-4 mb-6">
          <p className="text-center text-sm font-medium text-gray-700 mb-4">
            Pengaturan Tanggal & Nama Tempat <br /> untuk Rekap Semester pada
            File PDF
          </p>
          <div className="flex flex-col md:flex-row gap-4 items-center justify-center">
            <div className="text-center">
              <p className="text-sm text-gray-500 mb-2">Pilih Tanggal</p>
              <input
                type="date"
                value={selectedDate}
                onChange={(e) => setSelectedDate(e.target.value)}
                className="border border-gray-300 rounded-lg px-1 py-0.5 shadow-sm bg-white min-w-32"
              />
            </div>
          </div>
        </div>

        <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-6">
          <div className="bg-green-50 border border-green-200 rounded-lg p-3 text-center">
            <div className="text-green-600 font-bold text-lg">
              {statusSummary.Hadir}
            </div>
            <div className="text-green-700 text-sm">Hadir</div>
          </div>
          <div className="bg-yellow-50 border border-yellow-200 rounded-lg p-3 text-center">
            <div className="text-yellow-600 font-bold text-lg">
              {statusSummary.Izin}
            </div>
            <div className="text-yellow-700 text-sm">Izin</div>
          </div>
          <div className="bg-blue-50 border border-blue-200 rounded-lg p-3 text-center">
            <div className="text-blue-600 font-bold text-lg">
              {statusSummary.Sakit}
            </div>
            <div className="text-blue-700 text-sm">Sakit</div>
          </div>
          <div className="bg-red-50 border border-red-200 rounded-lg p-3 text-center">
            <div className="text-red-600 font-bold text-lg">
              {statusSummary.Alpha}
            </div>
            <div className="text-red-700 text-sm">Alpha</div>
          </div>
        </div>

        {loading ? (
          <div className="text-center py-8">
            <p className="text-gray-500">Memuat rekap...</p>
          </div>
        ) : filteredRecapData.length === 0 ? (
          <div className="text-center py-8">
            <p className="text-gray-500">
              Tidak ada data rekap untuk Semester {selectedSemester} kelas{" "}
              {selectedKelas}.
            </p>
            <p className="text-sm text-gray-400 mt-2">
              Coba pilih kelas atau semester lain.
            </p>
          </div>
        ) : (
          <>
            {/* Persentase Keseluruhan */}
            {(() => {
              const total =
                statusSummary.Hadir +
                statusSummary.Alpha +
                statusSummary.Izin +
                statusSummary.Sakit;
              const pctHadir =
                total > 0
                  ? ((statusSummary.Hadir / total) * 100).toFixed(1)
                  : "0";
              const pctAlpha =
                total > 0
                  ? ((statusSummary.Alpha / total) * 100).toFixed(1)
                  : "0";
              const pctIzin =
                total > 0
                  ? ((statusSummary.Izin / total) * 100).toFixed(1)
                  : "0";
              const pctSakit =
                total > 0
                  ? ((statusSummary.Sakit / total) * 100).toFixed(1)
                  : "0";
              return (
                <div className="mb-4 p-3 bg-gray-50 border border-gray-200 rounded-lg">
                  <p className="text-center text-sm font-semibold text-gray-700 mb-2">
                    Persentase Kehadiran Semester {selectedSemester} — Kelas{" "}
                    {selectedKelas}
                  </p>
                  <div className="grid grid-cols-4 gap-2 text-center text-sm">
                    <div>
                      <div className="text-green-600 font-bold text-base">
                        {pctHadir}%
                      </div>
                      <div className="text-gray-500">Hadir</div>
                    </div>
                    <div>
                      <div className="text-yellow-600 font-bold text-base">
                        {pctIzin}%
                      </div>
                      <div className="text-gray-500">Izin</div>
                    </div>
                    <div>
                      <div className="text-blue-600 font-bold text-base">
                        {pctSakit}%
                      </div>
                      <div className="text-gray-500">Sakit</div>
                    </div>
                    <div>
                      <div className="text-red-600 font-bold text-base">
                        {pctAlpha}%
                      </div>
                      <div className="text-gray-500">Alpha</div>
                    </div>
                  </div>
                </div>
              );
            })()}

            <div className="overflow-x-auto">
              <table className="min-w-full border-collapse border border-gray-200">
                <thead>
                  <tr className="bg-gray-100">
                    <th className="border border-gray-200 px-2 py-0.5 text-center text-sm font-semibold text-gray-700">
                      No.
                    </th>
                    <th className="border border-gray-200 px-1 py-0.5 text-left text-sm font-semibold text-gray-700">
                      Nama
                    </th>
                    <th className="border border-gray-200 px-1 py-0.5 text-left text-sm font-semibold text-gray-700">
                      Kelas
                    </th>
                    <th className="border border-gray-200 px-1 py-0.5 text-center text-sm font-semibold text-gray-700">
                      Hadir
                    </th>
                    <th className="border border-gray-200 px-1 py-0.5 text-center text-sm font-semibold text-gray-700">
                      Alpha
                    </th>
                    <th className="border border-gray-200 px-1 py-0.5 text-center text-sm font-semibold text-gray-700">
                      Izin
                    </th>
                    <th className="border border-gray-200 px-1 py-0.5 text-center text-sm font-semibold text-gray-700">
                      Sakit
                    </th>
                    <th className="border border-gray-200 px-1 py-0.5 text-center text-sm font-semibold text-gray-700">
                      % Hadir
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {filteredRecapData.map((item, index) => (
                    <tr
                      key={index}
                      className={index % 2 === 0 ? "bg-white" : "bg-gray-50"}
                    >
                      <td className="border border-gray-200 px-2 py-0.5 text-center text-sm text-gray-600 font-medium">
                        {index + 1}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-sm text-gray-600">
                        {item.nama || "N/A"}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-sm text-gray-600">
                        {item.kelas || "N/A"}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-center text-sm text-gray-600">
                        {item.hadir || 0}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-center text-sm text-gray-600">
                        {item.alpa || 0}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-center text-sm text-gray-600">
                        {item.izin || 0}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-center text-sm text-gray-600">
                        {item.sakit || 0}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-center text-sm text-gray-600">
                        {item.persenHadir !== undefined
                          ? `${item.persenHadir}%`
                          : "N/A"}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div className="mt-6 flex gap-4 justify-center">
              <button
                onClick={downloadExcel}
                className="px-1 py-0.5 bg-green-600 hover:bg-green-700 text-white rounded-lg font-medium transition-colors"
              >
                📥 Download Excel
              </button>
              <button
                onClick={downloadPDF}
                className="px-1 py-0.5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg font-medium transition-colors"
              >
                📄 Download PDF
              </button>
              {/* ✅ TOMBOL BARU */}
              <button
                onClick={() => setShowPreviewModal(true)}
                className="px-1 py-0.5 bg-purple-600 hover:bg-purple-700 text-white rounded-lg font-medium transition-colors"
              >
                👁️ Preview Pengurangan
              </button>
            </div>
          </>
        )}
      </div>
      {/* ✅ MODAL PREVIEW POTONGAN */}
      {showPreviewModal && (
        <div className="fixed inset-0 z-50 overflow-y-auto bg-black bg-opacity-50">
          <div className="flex min-h-full items-center justify-center p-4">
            <div
              className="bg-white rounded-xl shadow-2xl w-full max-w-2xl flex flex-col"
              style={{ maxHeight: "90vh" }}
            >
              {/* Header Modal */}
              <div className="flex items-center justify-between px-6 py-4 border-b border-gray-200">
                <div>
                  <h3 className="text-lg font-bold text-gray-800">
                    👁️ Preview Pengurangan Data Kehadiran
                  </h3>
                  <p className="text-xs text-gray-500 mt-0.5">
                    Kelas {selectedKelas} — Semester {selectedSemester}
                  </p>
                </div>
                <button
                  onClick={() => setShowPreviewModal(false)}
                  className="text-gray-400 hover:text-gray-600 text-2xl font-bold leading-none"
                >
                  ×
                </button>
              </div>

              {/* Keterangan aturan potongan */}
              <div className="px-6 py-3 bg-yellow-50 border-b border-yellow-100">
                <p className="text-xs font-semibold text-yellow-700 mb-1">
                  Aturan Pengurangan (berlaku untuk Alpha, Izin, dan Sakit
                  masing-masing):
                </p>
                <div className="flex flex-wrap gap-2 text-xs text-yellow-800">
                  {[
                    "0–3 → pengurangan 0",
                    "4–6 → pengurangan 3",
                    "7–15 → pengurangan 6",
                    "16–20 → pengurangan 10",
                    "≥21 → pengurangan 15",
                  ].map((rule) => (
                    <span
                      key={rule}
                      className="bg-yellow-100 border border-yellow-200 rounded px-2 py-0.5"
                    >
                      {rule}
                    </span>
                  ))}
                </div>
              </div>

              {/* Tabel */}
              <div className="overflow-y-auto overflow-x-auto flex-1 px-6 py-4 min-h-0">
                <table className="min-w-full border-collapse border border-gray-200 text-sm">
                  <thead>
                    <tr className="bg-gray-100">
                      <th className="border border-gray-200 px-2 py-1 text-center font-semibold text-gray-700">
                        No.
                      </th>
                      <th className="border border-gray-200 px-2 py-1 text-left font-semibold text-gray-700">
                        Nama
                      </th>
                      <th className="border border-gray-200 px-2 py-1 text-center font-semibold text-gray-700">
                        Alpha
                      </th>
                      <th className="border border-gray-200 px-2 py-1 text-center font-semibold text-yellow-700 bg-yellow-50">
                        Pengurangan Alpha
                      </th>
                      <th className="border border-gray-200 px-2 py-1 text-center font-semibold text-gray-700">
                        Izin
                      </th>
                      <th className="border border-gray-200 px-2 py-1 text-center font-semibold text-yellow-700 bg-yellow-50">
                        Pengurangan Izin
                      </th>
                      <th className="border border-gray-200 px-2 py-1 text-center font-semibold text-gray-700">
                        Sakit
                      </th>
                      <th className="border border-gray-200 px-2 py-1 text-center font-semibold text-yellow-700 bg-yellow-50">
                        Pengurangan Sakit
                      </th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredRecapData.map((item, index) => {
                      const alpha = item.alpa || 0;
                      const izin = item.izin || 0;
                      const sakit = item.sakit || 0;
                      const potongAlpha = getPotongan(alpha);
                      const potongIzin = getPotongan(izin);
                      const potongSakit = getPotongan(sakit);

                      return (
                        <tr
                          key={index}
                          className={
                            index % 2 === 0 ? "bg-white" : "bg-gray-50"
                          }
                        >
                          <td className="border border-gray-200 px-2 py-0.5 text-center text-gray-600">
                            {index + 1}
                          </td>
                          <td className="border border-gray-200 px-2 py-0.5 text-gray-700">
                            {item.nama || "N/A"}
                          </td>
                          <td className="border border-gray-200 px-2 py-0.5 text-center text-gray-600">
                            {alpha}
                          </td>
                          <td className="border border-gray-200 px-2 py-0.5 text-center font-medium text-yellow-700 bg-yellow-50">
                            {potongAlpha > 0 ? `${potongAlpha}` : "0"}
                          </td>
                          <td className="border border-gray-200 px-2 py-0.5 text-center text-gray-600">
                            {izin}
                          </td>
                          <td className="border border-gray-200 px-2 py-0.5 text-center font-medium text-yellow-700 bg-yellow-50">
                            {potongIzin > 0 ? `${potongIzin}` : "0"}
                          </td>
                          <td className="border border-gray-200 px-2 py-0.5 text-center text-gray-600">
                            {sakit}
                          </td>
                          <td className="border border-gray-200 px-2 py-0.5 text-center font-medium text-yellow-700 bg-yellow-50">
                            {potongSakit > 0 ? `${potongSakit}` : "0"}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>

              {/* Footer Modal */}
              <div className="px-6 py-3 border-t border-gray-200 flex justify-between items-center">
                <div className="flex gap-2">
                  <button
                    onClick={downloadPreviewExcel}
                    disabled={downloadingPreviewExcel}
                    className="px-4 py-1.5 bg-green-600 hover:bg-green-700 text-white rounded-lg font-medium transition-colors text-sm disabled:opacity-50"
                  >
                    {downloadingPreviewExcel
                      ? "⏳ Membuat file..."
                      : "📥 Download Excel"}
                  </button>
                  <button
                    onClick={downloadPreviewPDF}
                    className="px-4 py-1.5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg font-medium transition-colors text-sm"
                  >
                    📄 Download PDF
                  </button>
                </div>
                <button
                  onClick={() => setShowPreviewModal(false)}
                  className="px-4 py-1.5 bg-gray-200 hover:bg-gray-300 text-gray-700 rounded-lg font-medium transition-colors text-sm"
                >
                  Tutup
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

const ClearDataTab: React.FC<{
  onRefresh?: () => void;
  onDataCleared?: () => void;
}> = ({ onRefresh, onDataCleared }) => {
  const [isClearing, setIsClearing] = useState<boolean>(false);

  const clearAllLocalData = () => {
    // Hapus semua data dari localStorage
    const keysToRemove = [
      "students",
      "studentData",
      "dataSiswa",
      "siswaData",
      "studentList",
      "daftarSiswa",
    ];
    keysToRemove.forEach((key) => {
      localStorage.removeItem(key);
    });

    // Hapus semua data dari sessionStorage
    sessionStorage.clear();

    // Hapus semua data yang mungkin ada dengan prefix tertentu
    for (let i = localStorage.length - 1; i >= 0; i--) {
      const key = localStorage.key(i);
      if (
        key &&
        (key.includes("student") ||
          key.includes("siswa") ||
          key.includes("data"))
      ) {
        localStorage.removeItem(key);
      }
    }

    console.log("Semua data lokal berhasil dihapus");
  };

  const handleClearData = async () => {
    if (
      !window.confirm(
        "Yakin ingin menghapus semua data di sheet Absensi dan DataSiswa?\n\nHeader akan tetap dipertahankan. Data siswa di aplikasi juga akan dihapus. Tindakan ini tidak dapat dibatalkan."
      )
    ) {
      return;
    }

    setIsClearing(true);

    // Langsung hapus data lokal dulu
    clearAllLocalData();

    // Buat AbortController untuk timeout
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 30000); // 30 detik timeout

    try {
      console.log("Mengirim request ke:", endpoint);
      console.log("Payload:", {
        type: "deleteAllDataDataSiswanAbsensi",
        sheet: "both",
      });

      const response = await fetch(endpoint, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          type: "deleteAllDataDataSiswanAbsensi",
          sheet: "both",
        }),
        // Gunakan AbortController untuk timeout
        signal: controller.signal,
      });

      // Clear timeout jika request berhasil
      clearTimeout(timeoutId);

      console.log("Status respons:", response.status);

      if (!response.ok) {
        throw new Error(
          `Server error: ${response.status} ${response.statusText}`
        );
      }

      const jsonResponse = await response.json();
      console.log("Respons JSON:", jsonResponse);

      if (jsonResponse.success) {
        // Pastikan data lokal benar-benar terhapus
        clearAllLocalData();

        // Trigger callback untuk parent component
        if (onDataCleared) {
          onDataCleared();
        }

        // Trigger event untuk memberitahu komponen lain bahwa data telah dihapus
        window.dispatchEvent(new CustomEvent("dataCleared"));

        // Trigger refresh untuk memuat ulang data dari server
        if (onRefresh) {
          onRefresh();
        }

        // Trigger event global untuk refresh
        window.dispatchEvent(new CustomEvent("refreshData"));

        alert(
          `✅ ${jsonResponse.message}\n\nData siswa di aplikasi juga telah dihapus. Halaman akan dimuat ulang.`
        );

        // Reload halaman untuk memastikan semua komponen direset
        setTimeout(() => {
          window.location.reload();
        }, 1000);
      } else {
        throw new Error(jsonResponse.message || "Gagal menghapus data");
      }
    } catch (error) {
      console.error("Error saat menghapus data:", error);

      const errorMessage =
        error instanceof Error
          ? error.message
          : "Terjadi kesalahan tidak diketahui";

      // Penanganan spesifik untuk error CORS atau jaringan
      if (
        errorMessage.includes("Failed to fetch") ||
        errorMessage.toLowerCase().includes("cors") ||
        errorMessage.includes("NetworkError")
      ) {
        console.warn(
          "Mendeteksi masalah CORS atau jaringan, mencoba fallback..."
        );
        try {
          await fetch(endpoint, {
            method: "POST",
            mode: "no-cors", // Fallback untuk CORS
            headers: {
              "Content-Type": "application/json",
            },
            body: JSON.stringify({
              type: "deleteAllDataDataSiswanAbsensi",
              sheet: "both",
            }),
          });
          console.log("Fallback request dikirim (no-cors)");

          // Pastikan data lokal benar-benar terhapus untuk fallback
          clearAllLocalData();

          // Trigger callback untuk parent component
          if (onDataCleared) {
            onDataCleared();
          }

          // Trigger event untuk memberitahu komponen lain bahwa data telah dihapus
          window.dispatchEvent(new CustomEvent("dataCleared"));

          // Trigger refresh untuk memuat ulang data dari server
          if (onRefresh) {
            onRefresh();
          }

          // Trigger event global untuk refresh
          window.dispatchEvent(new CustomEvent("refreshData"));

          alert(
            "✅ Data berhasil dihapus (CORS fallback). Data siswa di aplikasi juga telah dihapus. Halaman akan dimuat ulang."
          );

          // Reload halaman untuk memastikan semua komponen direset
          setTimeout(() => {
            window.location.reload();
          }, 1000);
        } catch (fallbackError) {
          console.error("Fallback error:", fallbackError);
          alert(
            `❌ Gagal menghapus data: ${
              fallbackError instanceof Error
                ? fallbackError.message
                : "Unknown error"
            }. Periksa koneksi jaringan atau endpoint.`
          );
        }
      } else if (
        errorMessage.includes("aborted") ||
        errorMessage.includes("timeout")
      ) {
        alert("❌ Gagal menghapus data: Permintaan timeout. Coba lagi nanti.");
      } else {
        alert(`❌ Gagal menghapus data: ${errorMessage}. Detail di console.`);
      }
    } finally {
      // Pastikan timeout dibersihkan
      clearTimeout(timeoutId);
      setIsClearing(false);
    }
  };

  return (
    <div className="max-w-4xl mx-auto pb-20">
      <div className="bg-white p-6 rounded-lg shadow-md">
        <h2 className="text-2xl font-bold text-center text-red-700 mb-6">
          🗑️ Hapus Data
        </h2>
        <div className="bg-red-50 border border-red-200 rounded-lg p-4 mb-6">
          <p className="text-sm text-red-700 font-semibold mb-2">Peringatan:</p>
          <p className="text-sm text-red-600">
            Tindakan ini akan menghapus semua data di sheet Absensi dan
            DataSiswa (kecuali header), serta data siswa yang tersimpan di
            aplikasi. Tindakan ini tidak dapat dibatalkan.
          </p>
        </div>
        <div className="flex justify-center">
          <button
            onClick={handleClearData}
            disabled={isClearing}
            className={`px-6 py-2 rounded-lg font-medium text-white ${
              isClearing
                ? "bg-gray-400 cursor-not-allowed"
                : "bg-red-600 hover:bg-red-700"
            } transition-colors duration-200`}
          >
            {isClearing ? "Memproses..." : "🗑️ Hapus Semua Data"}
          </button>
        </div>
      </div>
    </div>
  );
};

// Komponen SplashScreen
const SplashScreen: React.FC = () => {
  return (
    <div className="fixed inset-0 bg-white flex flex-col items-center justify-center z-50">
      <style>
        {`
          @keyframes pulse {
            0% {
              transform: scale(1);
              opacity: 1;
            }
            50% {
              transform: scale(1.2);
              opacity: 0.7;
            }
            100% {
              transform: scale(1);
              opacity: 1;
            }
          }
          .animate-pulse-custom {
            animation: pulse 2s infinite;
          }
        `}
      </style>
      <img
        src="\images\logo_1.png"
        alt="Logo Aplikasi"
        className="w-52 h-70 mb-4 animate-pulse-custom" //Pengaturan ukuran logo
      />
      <p className="text-gray-800 text-lg font-semibold mt-6">
        Hardianto, S.Pd
      </p>
    </div>
  );
};

const JadwalMengajarTab: React.FC<{
  onRefresh: () => void;
}> = ({ onRefresh }) => {
  const [jadwalList, setJadwalList] = useState<JadwalMengajar[]>([]);
  const [kelasOptions, setKelasOptions] = useState<string[]>([]);
  const [selectedKelas, setSelectedKelas] = useState("");
  const [selectedHari, setSelectedHari] = useState<string[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [isSaving, setIsSaving] = useState<boolean>(false);
  const [isDeleting, setIsDeleting] = useState<boolean>(false);
  const [editingKelas, setEditingKelas] = useState<string | null>(null);

  const hariOptions = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];

  useEffect(() => {
    fetchJadwalMengajar();
    fetchKelasOptions();
  }, []);

  const fetchJadwalMengajar = () => {
    setLoading(true);
    fetch(`${endpoint}?action=jadwalMengajar`)
      .then((res) => {
        if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
        return res.json();
      })
      .then((data) => {
        if (data.success) {
          setJadwalList(data.data || []);
        } else {
          alert("❌ Gagal memuat data jadwal: " + data.message);
          setJadwalList([]);
        }
        setLoading(false);
      })
      .catch((error) => {
        console.error("Error fetch:", error);
        alert("❌ Gagal memuat data jadwal. Cek console untuk detail.");
        setLoading(false);
      });
  };

  const fetchKelasOptions = () => {
    fetch(`${endpoint}?action=kelasOptions`)
      .then((res) => {
        if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
        return res.json();
      })
      .then((data) => {
        if (data.success) {
          setKelasOptions(data.data || []);
        }
      })
      .catch((error) => {
        console.error("Error fetch kelas options:", error);
      });
  };

  const handleHariToggle = (hari: string) => {
    setSelectedHari((prev) =>
      prev.includes(hari) ? prev.filter((h) => h !== hari) : [...prev, hari]
    );
  };

  const handleSubmit = () => {
    if (!selectedKelas || selectedHari.length === 0) {
      alert("⚠️ Kelas dan minimal satu hari wajib dipilih!");
      return;
    }

    setIsSaving(true);

    const hariString = selectedHari.join(", ");

    fetch(endpoint, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        type: "jadwalMengajar",
        kelas: selectedKelas,
        hari: hariString,
      }),
    })
      .then(() => {
        alert("✅ Jadwal mengajar berhasil ditambahkan!");
        setSelectedKelas("");
        setSelectedHari([]);
        fetchJadwalMengajar();
        onRefresh();
        setIsSaving(false);
      })
      .catch(() => {
        alert("❌ Gagal menambahkan jadwal mengajar.");
        setIsSaving(false);
      });
  };

  const handleEdit = (jadwal: JadwalMengajar) => {
    setSelectedKelas(jadwal.kelas);
    setSelectedHari(jadwal.hari.split(", ").map((h) => h.trim()));
    setEditingKelas(jadwal.kelas);
  };

  const handleUpdate = () => {
    if (!selectedKelas || selectedHari.length === 0 || !editingKelas) {
      alert("⚠️ Kelas dan minimal satu hari wajib dipilih!");
      return;
    }

    setIsSaving(true);

    const hariString = selectedHari.join(", ");

    fetch(endpoint, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        type: "editJadwalMengajar",
        kelasLama: editingKelas,
        kelasBaru: selectedKelas,
        hari: hariString,
      }),
    })
      .then(() => {
        alert("✅ Jadwal mengajar berhasil diperbarui!");
        setSelectedKelas("");
        setSelectedHari([]);
        setEditingKelas(null);
        fetchJadwalMengajar();
        onRefresh();
        setIsSaving(false);
      })
      .catch(() => {
        alert("❌ Gagal memperbarui jadwal mengajar.");
        setIsSaving(false);
      });
  };

  const handleDelete = (kelas: string) => {
    if (!confirm(`Yakin ingin menghapus jadwal mengajar kelas: ${kelas}?`)) {
      return;
    }

    setIsDeleting(true);

    fetch(endpoint, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        type: "deleteJadwalMengajar",
        kelas: kelas,
      }),
    })
      .then(() => {
        alert("✅ Jadwal mengajar berhasil dihapus!");
        fetchJadwalMengajar();
        onRefresh();
        setIsDeleting(false);
      })
      .catch(() => {
        alert("❌ Gagal menghapus jadwal mengajar.");
        setIsDeleting(false);
      });
  };

  const handleCancel = () => {
    setSelectedKelas("");
    setSelectedHari([]);
    setEditingKelas(null);
  };

  if (loading) {
    return (
      <div className="text-center py-8">
        <p className="text-gray-500">Memuat data jadwal mengajar...</p>
      </div>
    );
  }

  return (
    <div className="max-w-4xl mx-auto" style={{ paddingBottom: "70px" }}>
      <div className="bg-white p-6 rounded-lg shadow-md mb-6">
        <h2 className="text-xl font-bold mb-4 text-center text-blue-600">
          {editingKelas !== null
            ? "Edit Jadwal Mengajar"
            : "Tambah Jadwal Mengajar"}
        </h2>
        <div className="grid grid-cols-1 gap-4 mb-4">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-2">
              Pilih Kelas
            </label>
            <select
              value={selectedKelas}
              onChange={(e) => setSelectedKelas(e.target.value)}
              className="w-full border border-gray-300 px-4 py-2 rounded-lg"
              disabled={isSaving}
            >
              <option value="">-- Pilih Kelas --</option>
              {kelasOptions.map((kelas) => (
                <option key={kelas} value={kelas}>
                  {kelas}
                </option>
              ))}
            </select>
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-2">
              Pilih Hari (bisa lebih dari satu)
            </label>
            <div className="grid grid-cols-2 md:grid-cols-3 gap-2">
              {hariOptions.map((hari) => (
                <label
                  key={hari}
                  className={`flex items-center gap-2 p-3 border rounded-lg cursor-pointer transition-colors ${
                    selectedHari.includes(hari)
                      ? "bg-blue-50 border-blue-500"
                      : "bg-white border-gray-300 hover:bg-gray-50"
                  }`}
                >
                  <input
                    type="checkbox"
                    checked={selectedHari.includes(hari)}
                    onChange={() => handleHariToggle(hari)}
                    disabled={isSaving}
                    className="h-4 w-4 text-blue-600 rounded"
                  />
                  <span className="text-sm text-gray-700">{hari}</span>
                </label>
              ))}
            </div>
          </div>
        </div>

        <div className="flex justify-center gap-4">
          {editingKelas !== null ? (
            <>
              <button
                onClick={handleUpdate}
                disabled={isSaving}
                className={`px-6 py-2 rounded-lg font-medium transition-colors ${
                  isSaving
                    ? "bg-blue-400 cursor-not-allowed"
                    : "bg-blue-600 hover:bg-blue-700"
                } text-white`}
              >
                {isSaving ? "⏳ Menyimpan..." : "💾 Update"}
              </button>
              <button
                onClick={handleCancel}
                disabled={isSaving}
                className="px-6 py-2 bg-gray-500 hover:bg-gray-600 text-white rounded-lg font-medium"
              >
                ❌ Batal
              </button>
            </>
          ) : (
            <button
              onClick={handleSubmit}
              disabled={isSaving}
              className={`px-6 py-2 rounded-lg font-medium transition-colors ${
                isSaving
                  ? "bg-blue-400 cursor-not-allowed"
                  : "bg-blue-600 hover:bg-blue-700"
              } text-white`}
            >
              {isSaving ? "⏳ Menyimpan..." : "➕ Tambah Jadwal Mengajar"}
            </button>
          )}
        </div>
      </div>

      <div className="bg-white p-6 rounded-lg shadow-md">
        <h3 className="text-lg font-semibold text-gray-700 mb-4">
          Daftar Jadwal Mengajar ({jadwalList.length})
        </h3>
        {jadwalList.length === 0 ? (
          <p className="text-center text-gray-500 py-8">
            Belum ada data jadwal mengajar.
          </p>
        ) : (
          <div className="space-y-3">
            {jadwalList.map((jadwal, index) => (
              <div
                key={index}
                className="flex justify-between items-center bg-gray-50 border border-gray-200 px-4 py-3 rounded-lg"
              >
                <div>
                  <p className="font-medium text-gray-800">
                    Kelas: {jadwal.kelas}
                  </p>
                  <p className="text-sm text-gray-600">Hari: {jadwal.hari}</p>
                </div>
                <div className="flex gap-2">
                  <button
                    onClick={() => handleEdit(jadwal)}
                    disabled={isSaving || isDeleting}
                    className={`text-xs px-3 py-1 rounded transition-colors ${
                      isSaving || isDeleting
                        ? "bg-yellow-400 cursor-not-allowed"
                        : "bg-yellow-500 hover:bg-yellow-600"
                    } text-white`}
                  >
                    ✏️ Edit
                  </button>
                  <button
                    onClick={() => handleDelete(jadwal.kelas)}
                    disabled={isSaving || isDeleting}
                    className={`text-xs px-3 py-1 rounded transition-colors ${
                      isSaving || isDeleting
                        ? "bg-red-400 cursor-not-allowed"
                        : "bg-red-500 hover:bg-red-600"
                    } text-white`}
                  >
                    🗑️ Hapus
                  </button>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
};

type UserRole = "guru" | "siswa";

const PASSWORDS: Record<UserRole, string> = {
  guru: "1234",
  siswa: "1111",
};

const LoginScreen: React.FC<{ onLogin: (role: UserRole) => void }> = ({
  onLogin,
}) => {
  const [username, setUsername] = useState<UserRole>("guru");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");

  const handleLogin = () => {
    if (password === PASSWORDS[username]) {
      setError("");
      onLogin(username);
    } else {
      setError("❌ Password salah!");
    }
  };

  return (
    <div className="fixed inset-0 bg-gray-100 flex items-center justify-center z-50 p-4">
      <div className="bg-white rounded-lg shadow-xl p-8 w-full max-w-sm">
        <div className="flex justify-center mb-4">
          <img src="\images\logo_2.png" alt="Logo" className="w-20 h-20" />
        </div>
        <h2 className="text-2xl font-bold text-center text-blue-700 mb-6">
          🔐 Login
        </h2>

        <label className="block text-sm font-medium text-gray-700 mb-1">
          Username
        </label>
        <select
          value={username}
          onChange={(e) => setUsername(e.target.value as UserRole)}
          className="w-full border border-gray-300 px-4 py-2 rounded-lg mb-4 bg-white"
        >
          <option value="guru">Guru</option>
          <option value="siswa">Siswa</option>
        </select>

        <label className="block text-sm font-medium text-gray-700 mb-1">
          Password
        </label>
        <input
          type="password"
          value={password}
          onChange={(e) => setPassword(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter") handleLogin();
          }}
          placeholder="Masukkan password"
          className="w-full border border-gray-300 px-4 py-2 rounded-lg mb-2"
        />

        {error && <p className="text-red-600 text-sm mb-2">{error}</p>}

        <button
          onClick={handleLogin}
          className="w-full mt-2 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-lg font-medium transition-colors"
        >
          Masuk
        </button>
      </div>
    </div>
  );
};

const StudentAttendanceApp: React.FC = () => {
  const [students, setStudents] = useState<Student[]>([]);
  const [uniqueClasses, setUniqueClasses] = useState<string[]>(["Semua"]);
  const [activeTab, setActiveTab] = useState<
    | "schoolData"
    | "studentData"
    | "attendance"
    | "recap"
    | "graph"
    | "semesterRecap"
    | "daftarHadir"
    | "tanggalMerah"
    | "jadwalMengajar"
    | "clearData"
  >("studentData");
  const [refreshTrigger, setRefreshTrigger] = useState(0);
  const [isSidebarOpen, setIsSidebarOpen] = useState(false);
  const [isLoading, setIsLoading] = useState(true);
  const [statusGuru, setStatusGuru] = useState<string | null>(null);
  const [loadingStudents, setLoadingStudents] = useState(true);
  const [userRole, setUserRole] = useState<UserRole | null>(() => {
    const saved = sessionStorage.getItem("userRole");
    return saved === "guru" || saved === "siswa" ? saved : null;
  });

  const handleLogin = (role: UserRole) => {
    sessionStorage.setItem("userRole", role);
    setUserRole(role);
    // Siswa langsung diarahkan ke halaman Absensi
    setActiveTab(role === "siswa" ? "attendance" : "studentData");
  };

  const handleLogout = () => {
    sessionStorage.removeItem("userRole");
    setUserRole(null);
    setIsSidebarOpen(false);
  };

  const applyStudents = (data: Student[]) => {
    setStudents(data);
    const classSet = new Set<string>();
    data.forEach((s) => {
      const k = s.kelas != null ? String(s.kelas).trim() : "";
      if (k && k !== "undefined" && k !== "null") classSet.add(k);
    });
    const classes = Array.from(classSet).sort((a, b) => {
      const aN = /^\d+$/.test(a),
        bN = /^\d+$/.test(b);
      if (aN && bN) return parseInt(a) - parseInt(b);
      if (aN) return -1;
      if (bN) return 1;
      return a.localeCompare(b);
    });
    setUniqueClasses(["Semua", ...classes]);
    setLoadingStudents(false);
  };

  const loadStudentsFromCache = async () => {
    const data = await getStudentsLocal();
    applyStudents(data as Student[]);
  };

  const fetchStudents = async () => {
    try {
      const changed = await syncAll();
      if (changed.length > 0) await loadStudentsFromCache();
    } catch (e) {
      console.error("Sync error:", e);
      setLoadingStudents(false);
    }
  };

  const fetchStatusGuru = () => {
    fetch(`${endpoint}?action=statusGuru`)
      .then((res) => {
        if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
        return res.json();
      })
      .then((data) => {
        if (data.success) {
          setStatusGuru(data.statusGuru);
        }
      })
      .catch((error) => {
        console.error("Error fetching status guru:", error);
      });
  };

  const handleRecapRefresh = () => {
    setRefreshTrigger((prev) => prev + 1);
  };

  const handleRefresh = () => {
    fetchStudents();
    fetchStatusGuru();
  };

  useEffect(() => {
    loadStudentsFromCache().then(() => fetchStudents());
    fetchStatusGuru();

    // Jika sudah login sebagai siswa (setelah refresh), buka tab Absensi
    if (sessionStorage.getItem("userRole") === "siswa") {
      setActiveTab("attendance");
    }

    const timer = setTimeout(() => {
      setIsLoading(false); // splash hanya tampilan
    }, 2000);

    return () => clearTimeout(timer);
  }, []);

  if (isLoading) {
    return <SplashScreen />;
  }

  // Belum login → tampilkan halaman login
  if (!userRole) {
    return <LoginScreen onLogin={handleLogin} />;
  }

  const isGuruKelas = statusGuru === "Guru Kelas";
  // Menu hanya tampil setelah status diketahui dan bukan Guru Kelas
  const shouldShowJadwalMengajar = statusGuru !== null && !isGuruKelas;

  return (
    <div className="min-h-screen bg-gray-100 flex flex-col">
      {/* Sidebar */}
      <aside
        className={`bg-white shadow-md w-64 space-y-2 py-6 px-2 fixed h-full top-0 left-0 transform ${
          isSidebarOpen ? "translate-x-0" : "-translate-x-full"
        } transition-transform duration-300 ease-in-out z-50`}
      >
        <div className="flex justify-between items-center mb-4 px-4">
          <h2 className="text-xl font-bold text-gray-800">Menu</h2>
          <button
            onClick={() => setIsSidebarOpen(false)}
            className="text-gray-600 hover:text-gray-800 text-2xl"
          >
            ✖️
          </button>
        </div>
        {[
          { tab: "schoolData", label: "🏫 Data Sekolah" },
          { tab: "studentData", label: "👥 Data Siswa" },
          { tab: "attendance", label: "📋 Absensi" },
          { tab: "recap", label: "📊 Rekap Bulanan" },
          { tab: "semesterRecap", label: "📚 Rekap Semester" },
          { tab: "graph", label: "📈 Grafik" },
          { tab: "daftarHadir", label: "📜 Riwayat Absen" },
          { tab: "tanggalMerah", label: "📅 Data Tanggal Merah" },
          ...(shouldShowJadwalMengajar
            ? [{ tab: "jadwalMengajar", label: "🗓️ Jadwal Mengajar" }]
            : []),
          { tab: "clearData", label: "🗑️ Hapus Data" },
        ]
          .filter(
            ({ tab }) =>
              userRole === "guru" ||
              ["attendance", "recap", "semesterRecap", "graph"].includes(tab)
          )
          .map(({ tab, label }) => (
            <button
              key={tab}
              onClick={() => {
                setActiveTab(
                  tab as
                    | "schoolData"
                    | "studentData"
                    | "attendance"
                    | "recap"
                    | "graph"
                    | "semesterRecap"
                    | "clearData"
                );
                setIsSidebarOpen(false);
              }}
              className={`w-full text-left py-2 px-4 rounded-lg text-sm font-medium transition-colors flex items-center gap-2 ${
                activeTab === tab
                  ? "bg-blue-600 text-white"
                  : "text-gray-600 hover:bg-gray-100"
              }`}
            >
              {label}
            </button>
          ))}

        <button
          onClick={handleLogout}
          className="w-full text-left py-2 px-4 rounded-lg text-sm font-medium text-red-600 hover:bg-red-50 transition-colors flex items-center gap-2 mt-4 border-t border-gray-200 pt-4"
        >
          🚪 Logout ({userRole === "guru" ? "Guru" : "Siswa"})
        </button>
      </aside>

      {/* Hamburger Menu Button */}
      <div className="fixed top-4 left-4 z-50">
        <button
          onClick={() => setIsSidebarOpen(!isSidebarOpen)}
          className="bg-blue-600 text-white px-4 py-2 rounded-lg shadow-md hover:bg-blue-700 transition-colors"
        >
          {isSidebarOpen ? "✖️ Tutup Menu" : "☰ Buka Menu"}
        </button>
      </div>

      {/* Logo di pojok kanan atas */}
      <div className="absolute top-4 right-4 z-50">
        <img
          src="\images\logo_2.png"
          alt="Logo Aplikasi"
          className="w-16 h-16"
        />
      </div>

      {/* Main Content */}
      <main
        className={`flex-1 p-6 transition-all duration-300 ${
          isSidebarOpen ? "ml-64" : "ml-0"
        } mt-16`}
      >
        <div className="text-center mb-6">
          <h1 className="text-3xl font-bold text-gray-800 mb-2">
            Sistem Absensi Siswa
          </h1>
          <p className="text-gray-600">Kelola data siswa dan absensi harian</p>
        </div>

        <div className="py-4">
          {userRole === "siswa" &&
            !["attendance", "recap", "semesterRecap", "graph"].includes(
              activeTab
            ) && (
              <div className="text-center text-gray-500 py-8">
                ⛔ Anda tidak memiliki akses ke halaman ini.
              </div>
            )}
          {activeTab === "schoolData" && userRole === "guru" && (
            <SchoolDataTab onRefresh={handleRefresh} />
          )}
          {activeTab === "studentData" && userRole === "guru" && (
            <StudentDataTab
              students={students}
              onRefresh={() => setTimeout(fetchStudents, 1500)}
              uniqueClasses={uniqueClasses}
              loading={loadingStudents}
            />
          )}
          {activeTab === "attendance" && (
            <AttendanceTab
              students={students}
              onRecapRefresh={handleRecapRefresh}
              statusGuru={statusGuru}
              userRole={userRole}
            />
          )}
          {activeTab === "recap" && (
            <MonthlyRecapTab
              onRefresh={handleRecapRefresh}
              uniqueClasses={uniqueClasses}
              students={students} // ✅ TAMBAHKAN INI
            />
          )}
          {activeTab === "graph" && <GraphTab uniqueClasses={uniqueClasses} />}
          {activeTab === "semesterRecap" && (
            <SemesterRecapTab
              uniqueClasses={uniqueClasses}
              students={students} // ✅ TAMBAHKAN INI
            />
          )}
          {activeTab === "daftarHadir" && userRole === "guru" && (
            <DaftarHadirTab students={students} uniqueClasses={uniqueClasses} />
          )}
          {activeTab === "tanggalMerah" && userRole === "guru" && (
            <TanggalMerahTab onRefresh={handleRefresh} />
          )}
          {activeTab === "jadwalMengajar" &&
            userRole === "guru" &&
            shouldShowJadwalMengajar && (
              <JadwalMengajarTab onRefresh={handleRefresh} />
            )}
          {activeTab === "clearData" && userRole === "guru" && <ClearDataTab />}
        </div>
      </main>
    </div>
  );
};

const hasAnyAttendanceOnDate = (
  students: Student[],
  day: number,
  attendanceDataGetter: (student: Student) => {
    attendance: { [day: number]: string };
  }
): boolean => {
  return students.some((student) => {
    const { attendance } = attendanceDataGetter(student);
    return attendance[day] && attendance[day] !== "";
  });
};

const DaftarHadirTab: React.FC<{
  students: Student[];
  uniqueClasses: string[];
}> = ({ students, uniqueClasses }) => {
  const [attendanceData, setAttendanceData] = useState<AttendanceHistory[]>([]);
  const [selectedKelas, setSelectedKelas] = useState<string>("Semua");
  const [selectedMonth, setSelectedMonth] = useState<number>(
    new Date().getMonth() + 1
  ); // 1-12
  const [selectedYear, setSelectedYear] = useState<number>(
    new Date().getFullYear()
  );
  const [loading, setLoading] = useState<boolean>(true);
  const [editedRecords, setEditedRecords] = useState<
    Record<string, EditedRecord>
  >({});
  const [isSaving, setIsSaving] = useState<boolean>(false);
  const [deletingStudentId, setDeletingStudentId] = useState<string | null>(
    null
  );
  const [schoolData, setSchoolData] = useState<SchoolData | null>(null);
  const [selectedDate, setSelectedDate] = useState<string>(
    new Date().toISOString().split("T")[0]
  );

  const [selectedStudent, setSelectedStudent] = useState<Student | null>(null);
  const [showModal, setShowModal] = useState<boolean>(false);
  const [isDeleting, setIsDeleting] = useState<boolean>(false);
  const [tanggalMerahList, setTanggalMerahList] = useState<TanggalMerah[]>([]);
  const [loadingTanggalMerah, setLoadingTanggalMerah] =
    useState<boolean>(false);
  const [sundayDays, setSundayDays] = useState<Set<number>>(new Set());
  const [filteredTanggalMerah, setFilteredTanggalMerah] = useState<
    TanggalMerah[]
  >([]);
  const [jadwalMengajar, setJadwalMengajar] = useState<JadwalMengajar[]>([]);
  const [loadingJadwal, setLoadingJadwal] = useState<boolean>(false);

  const [customColors, setCustomColors] = useState({
    hariMinggu: "#DC3545", // Merah default
    liburSemester: "#22C55E", // Hijau default
    tanggalMerah: "#D3D3D3", // Abu-abu default
    bukanJadwal: "#93C5FD", // Biru muda default
    jadwalMengajar: "#FFFFFF", // Putih default
    dataEdit: "#FEF3C7", // Kuning default
  });

  const [showColorPicker, setShowColorPicker] = useState(false);

  useEffect(() => {
    const savedColors = localStorage.getItem("daftarHadirColors");
    if (savedColors) {
      try {
        setCustomColors(JSON.parse(savedColors));
      } catch (error) {
        console.error("Error loading colors:", error);
      }
    }
  }, []);

  // Simpan warna ke localStorage setiap kali berubah
  const handleColorChange = (colorKey: string, newColor: string) => {
    const updatedColors = { ...customColors, [colorKey]: newColor };
    setCustomColors(updatedColors);
    localStorage.setItem("daftarHadirColors", JSON.stringify(updatedColors));
  };

  // Reset ke warna default
  const resetColors = () => {
    const defaultColors = {
      hariMinggu: "#DC3545",
      liburSemester: "#22C55E",
      tanggalMerah: "#D3D3D3",
      bukanJadwal: "#93C5FD",
      jadwalMengajar: "#FFFFFF",
      dataEdit: "#FEF3C7",
    };
    setCustomColors(defaultColors);
    localStorage.setItem("daftarHadirColors", JSON.stringify(defaultColors));
    alert("✅ Warna berhasil direset ke default!");
  };

  // Fungsi untuk mengecek apakah tanggal adalah hari Minggu
  const isSunday = (day: number): boolean => {
    const date = new Date(selectedYear, selectedMonth - 1, day);
    return date.getDay() === 0; // 0 = Sunday
  };

  const months = [
    { value: 1, label: "Januari" },
    { value: 2, label: "Februari" },
    { value: 3, label: "Maret" },
    { value: 4, label: "April" },
    { value: 5, label: "Mei" },
    { value: 6, label: "Juni" },
    { value: 7, label: "Juli" },
    { value: 8, label: "Agustus" },
    { value: 9, label: "September" },
    { value: 10, label: "Oktober" },
    { value: 11, label: "November" },
    { value: 12, label: "Desember" },
  ];

  const years = Array.from({ length: 11 }, (_, i) => 2020 + i); // 2020-2030, sesuaikan jika perlu

  // Hitung jumlah hari di bulan (pindahkan ke atas agar bisa digunakan di useMemo)
  const daysInMonth = new Date(selectedYear, selectedMonth, 0).getDate();

  // Filter siswa berdasarkan kelas (pindahkan ke atas sebelum memoization)
  const filteredStudents = React.useMemo(() => {
    const result =
      selectedKelas === "Semua"
        ? students
        : students.filter(
            (student) => String(student.kelas).trim() === selectedKelas
          );

    console.log("=== FILTERED STUDENTS ===");
    console.log("Selected Kelas:", selectedKelas);
    console.log("Total students:", result.length);
    if (result.length > 0) {
      console.log("Sample students (first 3):");
      result.slice(0, 3).forEach((s, i) => {
        console.log(`Student ${i}:`, {
          name: s.name,
          nisn: s.nisn,
          kelas: s.kelas,
        });
      });
    }

    return result;
  }, [students, selectedKelas]);

  // Define getAttendanceForStudent FIRST before using in useMemo
  const getAttendanceForStudent = React.useCallback(
    (student: Student) => {
      const studentAttendance: Record<number, string> = {};

      const studentNisn = String(student.nisn || "")
        .trim()
        .replace(/\s+/g, "")
        .toUpperCase();
      const studentNama = String(student.name || "")
        .trim()
        .toLowerCase();

      attendanceData.forEach((record) => {
        const recordNisn = String(record.nisn || "")
          .trim()
          .replace(/\s+/g, "")
          .toUpperCase();
        const recordNama = String(record.nama || "")
          .trim()
          .toLowerCase();

        const nisnMatch =
          studentNisn && recordNisn && studentNisn === recordNisn;
        const nameMatch =
          studentNama && recordNama && studentNama === recordNama;
        const isMatch = nisnMatch || nameMatch;

        if (isMatch) {
          const dateParts = record.tanggal.split("/");
          if (dateParts.length === 3) {
            const day = parseInt(dateParts[0], 10);
            const month = parseInt(dateParts[1], 10);
            const year = parseInt(dateParts[2], 10);

            if (
              month === selectedMonth &&
              year === selectedYear &&
              !isNaN(day)
            ) {
              let code = "";
              switch (record.status) {
                case "Hadir":
                  code = "H";
                  break;
                case "Izin":
                  code = "I";
                  break;
                case "Sakit":
                  code = "S";
                  break;
                case "Alpha":
                  code = "A";
                  break;
              }
              if (code) studentAttendance[day] = code;
            }
          }
        }
      });

      Object.entries(editedRecords).forEach(([key, record]) => {
        const keyParts = key.split("_");
        if (keyParts.length >= 2 && keyParts[0] === student.id) {
          const day = parseInt(keyParts[1], 10);
          const dateParts = record.date.split("/");
          if (dateParts.length === 3) {
            const month = parseInt(dateParts[1], 10);
            const year = parseInt(dateParts[2], 10);

            if (
              month === selectedMonth &&
              year === selectedYear &&
              !isNaN(day)
            ) {
              let code = "";
              switch (record.status) {
                case "Hadir":
                  code = "H";
                  break;
                case "Izin":
                  code = "I";
                  break;
                case "Sakit":
                  code = "S";
                  break;
                case "Alpha":
                  code = "A";
                  break;
              }
              if (code) {
                studentAttendance[day] = code;
              } else if (record.status === "") {
                delete studentAttendance[day];
              }
            }
          }
        }
      });

      const countH = Object.values(studentAttendance).filter(
        (v) => v === "H"
      ).length;
      const countS = Object.values(studentAttendance).filter(
        (v) => v === "S"
      ).length;
      const countI = Object.values(studentAttendance).filter(
        (v) => v === "I"
      ).length;
      const countA = Object.values(studentAttendance).filter(
        (v) => v === "A"
      ).length;

      return {
        attendance: studentAttendance,
        counts: { H: countH, S: countS, I: countI, A: countA },
      };
    },
    [attendanceData, editedRecords, selectedMonth, selectedYear]
  );

  // NOW use it in useMemo
  const studentAttendanceMap = React.useMemo(() => {
    const map = new Map();
    filteredStudents.forEach((student) => {
      map.set(student.id, getAttendanceForStudent(student));
    });
    return map;
  }, [filteredStudents, getAttendanceForStudent]);

  const attendanceByDateMemo = React.useMemo(() => {
    const attendanceByDate: {
      [day: number]: { hadir: number; total: number };
    } = {};

    filteredStudents.forEach((student) => {
      const cached = studentAttendanceMap.get(student.id);
      if (!cached) return;

      const { attendance } = cached;

      Array.from({ length: daysInMonth }, (_, day) => {
        const dayNum = day + 1;
        const status = attendance[dayNum] || "";

        if (!attendanceByDate[dayNum]) {
          attendanceByDate[dayNum] = { hadir: 0, total: 0 };
        }

        if (status === "H") {
          attendanceByDate[dayNum].hadir += 1;
        }

        if (status !== "") {
          attendanceByDate[dayNum].total += 1;
        }
      });
    });

    return attendanceByDate;
  }, [filteredStudents, studentAttendanceMap, daysInMonth]);

  const daysWithNoData = React.useMemo(() => {
    const days = new Set<number>();
    for (let day = 1; day <= daysInMonth; day++) {
      const stats = attendanceByDateMemo[day] || { hadir: 0, total: 0 };
      if (stats.total === 0) {
        days.add(day);
      }
    }
    return days;
  }, [attendanceByDateMemo, daysInMonth]);

  const fetchAttendanceData = async () => {
    setLoading(true);
    try {
      const res = await fetch(`${endpoint}?action=attendanceHistory`);
      if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
      const data = await res.json();
      if (data.success) {
        console.log("=== RAW DATA FROM SERVER ===");
        console.log("Total records:", data.data.length);
        // Log sample data (tetap jika ingin debug)
        if (data.data.length > 0) {
          console.log("Sample records (first 5):");
          data.data.slice(0, 5).forEach((record: any, index: number) => {
            console.log(`Record ${index}:`, {
              tanggal: record.tanggal,
              nama: record.nama,
              kelas: record.kelas,
              nisn: record.nisn,
              status: record.status,
            });
          });
        }
        const validData = filterValidAttendance(data.data || []);
        console.log("Valid records after filter:", validData.length);
        // Cek unique NISN dan dates (tetap jika ingin debug)
        const uniqueNISN = new Set(
          validData.map((r) => String(r.nisn || "").trim())
        );
        console.log("Unique NISN in attendance data:", Array.from(uniqueNISN));
        const uniqueDates = new Set(validData.map((r) => r.tanggal));
        console.log("Unique dates:", Array.from(uniqueDates).sort());
        setAttendanceData(validData);
      } else {
        alert("❌ Gagal memuat data absensi: " + data.message);
        setAttendanceData([]);
      }
    } catch (error) {
      console.error("Error fetch:", error);
      alert("❌ Gagal memuat data absensi. Cek console untuk detail.");
    } finally {
      setLoading(false);
    }
  };

  // Hitung hari Minggu di bulan terpilih
  useEffect(() => {
    const sundays = new Set<number>();
    for (let day = 1; day <= daysInMonth; day++) {
      if (isSunday(day)) {
        sundays.add(day);
      }
    }
    setSundayDays(sundays);
  }, [selectedMonth, selectedYear, daysInMonth]);

  useEffect(() => {
    fetch(`${endpoint}?action=schoolData`)
      .then((res) => {
        if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
        return res.json();
      })
      .then((data) => {
        if (data.success && data.data && data.data.length > 0) {
          setSchoolData(data.data[0]);
        } else {
          setSchoolData(null);
        }
      })
      .catch((error) => {
        console.error("Error fetching school data:", error);
        alert("❌ Gagal memuat data sekolah. Cek console untuk detail.");
      });
  }, []); // Fetch sekali saat mount

  useEffect(() => {
    setLoading(true);
    fetch(`${endpoint}?action=attendanceHistory`)
      .then((res) => {
        if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
        return res.json();
      })
      .then((data) => {
        if (data.success) {
          console.log("=== RAW DATA FROM SERVER ===");
          console.log("Total records:", data.data.length);

          // Log sample data
          if (data.data.length > 0) {
            console.log("Sample records (first 5):");
            data.data.slice(0, 5).forEach((record: any, index: number) => {
              console.log(`Record ${index}:`, {
                tanggal: record.tanggal,
                nama: record.nama,
                kelas: record.kelas,
                nisn: record.nisn,
                status: record.status,
              });
            });
          }

          // Filter data yang valid
          const validData = filterValidAttendance(data.data || []);
          console.log("Valid records after filter:", validData.length);

          // Cek unique NISN dalam data
          const uniqueNISN = new Set(
            validData.map((r) => String(r.nisn || "").trim())
          );
          console.log(
            "Unique NISN in attendance data:",
            Array.from(uniqueNISN)
          );

          // Cek unique tanggal
          const uniqueDates = new Set(validData.map((r) => r.tanggal));
          console.log("Unique dates:", Array.from(uniqueDates).sort());

          setAttendanceData(validData);
        } else {
          alert("❌ Gagal memuat data absensi: " + data.message);
          setAttendanceData([]);
        }
        setLoading(false);
      })
      .catch((error) => {
        console.error("Error fetch:", error);
        alert("❌ Gagal memuat data absensi. Cek console untuk detail.");
        setLoading(false);
      });
  }, []);

  // Fetch tanggal merah
  useEffect(() => {
    fetchTanggalMerah();
  }, []);

  const fetchTanggalMerah = async () => {
    setLoadingTanggalMerah(true);
    try {
      const res = await fetch(`${endpoint}?action=tanggalMerah`);
      if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
      const data = await res.json();
      if (data.success) {
        console.log("Tanggal merah loaded:", data.data);
        setTanggalMerahList(data.data || []);
      } else {
        console.error("Gagal memuat tanggal merah:", data.message);
        setTanggalMerahList([]);
      }
    } catch (error) {
      console.error("Error fetch tanggal merah:", error);
    } finally {
      setLoadingTanggalMerah(false);
    }
  };

  // Filter tanggal merah berdasarkan bulan dan tahun yang dipilih
  useEffect(() => {
    const filtered = tanggalMerahList.filter((tm) => {
      const [dayStart, monthStart, yearStart] = tm.tanggal
        .split("/")
        .map(Number);

      // Jika tidak ada tanggal akhir, cek hanya tanggal mulai
      if (!tm.tanggalAkhir) {
        return monthStart === selectedMonth && yearStart === selectedYear;
      }

      // Jika ada tanggal akhir, cek apakah rentang mencakup bulan yang dipilih
      const [dayEnd, monthEnd, yearEnd] = tm.tanggalAkhir
        .split("/")
        .map(Number);

      const startDate = new Date(yearStart, monthStart - 1, dayStart);
      const endDate = new Date(yearEnd, monthEnd - 1, dayEnd);
      const currentMonthStart = new Date(selectedYear, selectedMonth - 1, 1);
      const currentMonthEnd = new Date(selectedYear, selectedMonth, 0); // Last day of month

      // Cek apakah rentang tanggal merah overlap dengan bulan yang dipilih
      return startDate <= currentMonthEnd && endDate >= currentMonthStart;
    });
    setFilteredTanggalMerah(filtered);
  }, [tanggalMerahList, selectedMonth, selectedYear]);

  const fetchJadwalMengajar = async () => {
    setLoadingJadwal(true);
    try {
      const res = await fetch(`${endpoint}?action=jadwalMengajar`);
      if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
      const data = await res.json();
      if (data.success) {
        console.log("Jadwal mengajar loaded:", data.data);
        setJadwalMengajar(data.data || []);
      } else {
        console.error("Gagal memuat jadwal mengajar:", data.message);
        setJadwalMengajar([]);
      }
    } catch (error) {
      console.error("Error fetch jadwal mengajar:", error);
    } finally {
      setLoadingJadwal(false);
    }
  };

  const isJadwalMengajar = (day: number): boolean => {
    // Jika guru kelas, semua hari adalah jadwal mengajar
    if (schoolData?.statusGuru === "Guru Kelas") {
      return true;
    }

    // Jika "Semua" dipilih, tidak bisa tentukan jadwal spesifik
    if (selectedKelas === "Semua") {
      return false; // ← UBAH dari true ke false
    }

    // Cari jadwal untuk kelas yang dipilih
    const jadwal = jadwalMengajar.find((j) => j.kelas === selectedKelas);

    if (!jadwal) {
      // Jika tidak ada jadwal untuk kelas ini, anggap semua hari bukan jadwal
      return false;
    }

    // Dapatkan nama hari dari tanggal
    const date = new Date(selectedYear, selectedMonth - 1, day);
    const dayNames = [
      "Minggu",
      "Senin",
      "Selasa",
      "Rabu",
      "Kamis",
      "Jumat",
      "Sabtu",
    ];
    const dayName = dayNames[date.getDay()];

    // Split hari dari jadwal dan trim setiap item
    const hariJadwal = jadwal.hari
      .split(",")
      .map((h) => h.trim())
      .filter((h) => h.length > 0);

    // Debug log
    console.log(`Day ${day} (${dayName}):`, {
      jadwalHari: hariJadwal,
      isMatch: hariJadwal.includes(dayName),
    });

    return hariJadwal.includes(dayName);
  };

  const isBukanJadwalMengajar = (day: number): boolean => {
    // Jika guru kelas, tidak ada yang bukan jadwal
    if (schoolData?.statusGuru === "Guru Kelas") {
      return false;
    }

    // Jika "Semua" dipilih, tidak tampilkan sebagai bukan jadwal
    if (selectedKelas === "Semua") {
      return false;
    }

    const isSundayDay = isSunday(day);
    const isTglMerah = isTanggalMerah(day);
    const isLiburSem = isLiburSemester(day);
    const isJadwal = isJadwalMengajar(day);

    // ✅ PERBAIKAN: Bukan jadwal mengajar jika: bukan minggu, bukan tanggal merah,
    // bukan libur semester, DAN bukan jadwal
    return !isSundayDay && !isTglMerah && !isLiburSem && !isJadwal;
  };

  // Fetch jadwal mengajar - PRIORITAS TINGGI
  useEffect(() => {
    fetchJadwalMengajar();
  }, []); // Fetch sekali saat mount

  // Fetch attendance data - setelah jadwal ter-load
  useEffect(() => {
    if (loadingJadwal) return; // Tunggu jadwal selesai di-load
    fetchAttendanceData();
  }, [loadingJadwal]); // Dependency: tunggu loading jadwal selesai

  const isInDateRange = (
    day: number,
    startDate: string,
    endDate?: string
  ): boolean => {
    const currentDate = `${String(day).padStart(2, "0")}/${String(
      selectedMonth
    ).padStart(2, "0")}/${selectedYear}`;

    if (!endDate) {
      return currentDate === startDate;
    }

    // Parse dates
    const [d1, m1, y1] = startDate.split("/").map(Number);
    const [d2, m2, y2] = endDate.split("/").map(Number);
    const [dC, mC, yC] = currentDate.split("/").map(Number);

    const start = new Date(y1, m1 - 1, d1);
    const end = new Date(y2, m2 - 1, d2);
    const current = new Date(yC, mC - 1, dC);

    return current >= start && current <= end;
  };

  // Helper function to check if a date is a tanggal merah
  const isTanggalMerah = (day: number): boolean => {
    return tanggalMerahList.some((tm) => {
      const isNotLiburSemester = !(
        tm.deskripsi.toLowerCase().includes("libur akhir semester") ||
        tm.deskripsi.toLowerCase().includes("libur semester")
      );

      if (!isNotLiburSemester) return false;

      return isInDateRange(day, tm.tanggal, tm.tanggalAkhir);
    });
  };

  // Helper function to check if a date is libur semester
  const isLiburSemester = (day: number): boolean => {
    return tanggalMerahList.some((tm) => {
      const isLibur =
        tm.deskripsi.toLowerCase().includes("libur akhir semester") ||
        tm.deskripsi.toLowerCase().includes("libur semester");

      if (!isLibur) return false;

      return isInDateRange(day, tm.tanggal, tm.tanggalAkhir);
    });
  };

  // Helper function to get deskripsi for libur semester
  const getLiburSemesterDeskripsi = (day: number): string => {
    const dateStr = `${String(day).padStart(2, "0")}/${String(
      selectedMonth
    ).padStart(2, "0")}/${selectedYear}`;

    const found = tanggalMerahList.find((tm) => tm.tanggal === dateStr);
    return found ? found.deskripsi : "";
  };

  // Helper function to get deskripsi for a tanggal merah
  const getTanggalMerahDeskripsi = (day: number): string => {
    const dateStr = `${String(day).padStart(2, "0")}/${String(
      selectedMonth
    ).padStart(2, "0")}/${selectedYear}`;

    const found = tanggalMerahList.find((tm) => tm.tanggal === dateStr);
    return found ? found.deskripsi : "";
  };

  // Filter data absensi yang valid (tidak ada formula atau error)
  const filterValidAttendance = (
    data: AttendanceHistory[]
  ): AttendanceHistory[] => {
    return data.filter((record) => {
      const hasValidTanggal =
        record.tanggal &&
        !record.tanggal.toString().startsWith("=") &&
        /^\d{2}\/\d{2}\/\d{4}$/.test(record.tanggal.toString());

      const hasValidNama =
        record.nama &&
        !record.nama.toString().startsWith("=") &&
        record.nama.toString().trim() !== "";

      const hasValidNisn =
        record.nisn &&
        !record.nisn.toString().startsWith("=") &&
        record.nisn.toString().trim() !== "";

      const hasValidStatus =
        record.status &&
        ["Hadir", "Izin", "Sakit", "Alpha"].includes(record.status.toString());

      return hasValidTanggal && hasValidNama && hasValidNisn && hasValidStatus;
    });
  };

  const getAttendanceByDate = () => {
    return attendanceByDateMemo; // Gunakan yang sudah di-memoize
  };

  const getTotalSummary = () => {
    let totalH = 0,
      totalS = 0,
      totalI = 0,
      totalA = 0;

    filteredStudents.forEach((student) => {
      const { counts } = getAttendanceForStudent(student);
      totalH += counts.H;
      totalS += counts.S;
      totalI += counts.I;
      totalA += counts.A;
    });

    const grandTotal = totalH + totalS + totalI + totalA;

    return {
      totalH,
      totalS,
      totalI,
      totalA,
      grandTotal,
      percentH: grandTotal > 0 ? ((totalH / grandTotal) * 100).toFixed(0) : "0",
      percentS: grandTotal > 0 ? ((totalS / grandTotal) * 100).toFixed(0) : "0",
      percentI: grandTotal > 0 ? ((totalI / grandTotal) * 100).toFixed(0) : "0",
      percentA: grandTotal > 0 ? ((totalA / grandTotal) * 100).toFixed(0) : "0",
    };
  };

  const getGenderSummary = () => {
    let totalLakiLaki = 0;
    let totalPerempuan = 0;

    filteredStudents.forEach((student) => {
      const jenisKelamin = String(student.jenisKelamin || "")
        .trim()
        .toUpperCase();
      if (jenisKelamin === "L" || jenisKelamin === "LAKI-LAKI") {
        totalLakiLaki++;
      } else if (jenisKelamin === "P" || jenisKelamin === "PEREMPUAN") {
        totalPerempuan++;
      }
    });

    return {
      lakiLaki: totalLakiLaki,
      perempuan: totalPerempuan,
      total: totalLakiLaki + totalPerempuan,
    };
  };

  const getHariEfektif = () => {
    let hariEfektif = 0;

    for (let day = 1; day <= daysInMonth; day++) {
      const isSundayDay = isSunday(day);
      const isTglMerah = isTanggalMerah(day);
      const isLiburSem = isLiburSemester(day);
      const isBukanJadwal = isBukanJadwalMengajar(day);

      // Hitung hanya hari yang tidak merah, minggu, atau libur semester
      if (!isSundayDay && !isTglMerah && !isLiburSem && !isBukanJadwal) {
        hariEfektif++;
      }
    }

    return hariEfektif;
  };

  const handleStatusChange = (
    student: Student,
    day: number,
    newStatus: AttendanceStatus | ""
  ) => {
    console.log(`\n=== Status Change ===`);
    console.log("Student:", student.name);
    console.log("Day:", day);
    console.log("New Status:", newStatus);
    console.log("Current editedRecords:", editedRecords);

    const key = `${student.id}_${day}`;
    const dateStr = `${String(day).padStart(2, "0")}/${String(
      selectedMonth
    ).padStart(2, "0")}/${selectedYear}`;

    if (newStatus === "") {
      // Cek apakah ada data asli dari attendanceData
      const hasOriginalData = attendanceData.some((record) => {
        const studentNisn = String(student.nisn || "")
          .trim()
          .replace(/\s+/g, "")
          .toUpperCase();
        const studentNama = String(student.name || "")
          .trim()
          .toLowerCase();
        const recordNisn = String(record.nisn || "")
          .trim()
          .replace(/\s+/g, "")
          .toUpperCase();
        const recordNama = String(record.nama || "")
          .trim()
          .toLowerCase();

        const isMatch =
          (studentNisn && recordNisn && studentNisn === recordNisn) ||
          (studentNama && recordNama && studentNama === recordNama);

        if (isMatch) {
          const dateParts = record.tanggal.split("/");
          if (dateParts.length === 3) {
            const recordDay = parseInt(dateParts[0], 10);
            const recordMonth = parseInt(dateParts[1], 10);
            const recordYear = parseInt(dateParts[2], 10);
            return (
              recordDay === day &&
              recordMonth === selectedMonth &&
              recordYear === selectedYear
            );
          }
        }
        return false;
      });

      if (hasOriginalData) {
        // Jika ada data asli, tandai untuk dihapus dengan status kosong
        setEditedRecords((prev) => ({
          ...prev,
          [key]: {
            date: dateStr,
            nisn: String(student.nisn || ""),
            status: "", // Status kosong = hapus
          },
        }));
      } else {
        // Jika tidak ada data asli, hapus dari editedRecords
        setEditedRecords((prev) => {
          const newRecords = { ...prev };
          delete newRecords[key];
          console.log("Removed from editedRecords. New state:", newRecords);
          return newRecords;
        });
      }
    } else {
      // Tambah/update di editedRecords
      const newRecord = {
        date: dateStr,
        nisn: String(student.nisn || ""),
        status: newStatus,
      };

      setEditedRecords((prev) => {
        const newRecords = {
          ...prev,
          [key]: newRecord,
        };
        console.log("Updated editedRecords. New state:", newRecords);
        return newRecords;
      });
    }
  };

  const handleSaveChanges = () => {
    if (Object.keys(editedRecords).length === 0) {
      alert("⚠️ Tidak ada perubahan untuk disimpan.");
      return;
    }

    setIsSaving(true);
    const updates = Object.values(editedRecords).map((record) => ({
      tanggal: record.date,
      nisn: record.nisn,
      status: record.status,
    }));

    fetch(endpoint, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        type: "bulkUpdateAttendance",
        updates,
      }),
    })
      .then(() => {
        alert("✅ Perubahan berhasil disimpan!");
        setEditedRecords({});
        // Update data absensi secara dinamis tanpa reload halaman
        fetchAttendanceData();
      })
      .catch(() => alert("❌ Gagal menyimpan perubahan."))
      .finally(() => setIsSaving(false));
  };

  const handleDeleteStudentAttendance = async (student: Student) => {
    const confirmMessage = `⚠️ PERINGATAN!\n\nAnda akan menghapus SEMUA riwayat absensi untuk:\n\nNama: ${
      student.name
    }\nNISN: ${student.nisn}\nKelas: ${
      student.kelas
    }\n\nData yang akan dihapus:\n- Semua riwayat kehadiran di bulan ${
      months.find((m) => m.value === selectedMonth)?.label
    } ${selectedYear}\n- Data di sheet "Absensi" di Google Sheets\n\nTindakan ini TIDAK DAPAT DIBATALKAN!\n\nApakah Anda yakin?`;

    if (!confirm(confirmMessage)) {
      return;
    }

    setDeletingStudentId(student.id);

    try {
      const response = await fetch(endpoint, {
        method: "POST",
        mode: "no-cors",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          type: "deleteStudentAttendanceByName",
          nama: student.name,
          bulan: selectedMonth,
          tahun: selectedYear,
        }),
      });

      alert(`✅ Riwayat absensi ${student.name} berhasil dihapus!`);

      // Hapus data dari state lokal
      setAttendanceData((prev) =>
        prev.filter((record) => record.nama !== student.name)
      );

      // Hapus editedRecords untuk siswa ini
      setEditedRecords((prev) => {
        const newRecords = { ...prev };
        Object.keys(newRecords).forEach((key) => {
          if (key.startsWith(`${student.id}_`)) {
            delete newRecords[key];
          }
        });
        return newRecords;
      });

      // Refresh data dari server
      fetchAttendanceData();
    } catch (error) {
      console.error("Error deleting student attendance:", error);
      alert("❌ Gagal menghapus riwayat absensi. Silakan coba lagi.");
    } finally {
      setDeletingStudentId(null);
    }
  };

  const handleDeleteAllAttendance = () => {
    const monthLabel =
      months.find((m) => m.value === selectedMonth)?.label || "";
    const kelasLabel =
      selectedKelas === "Semua" ? "semua kelas" : `kelas ${selectedKelas}`;

    if (
      confirm(
        `Yakin ingin menghapus semua data absensi ${kelasLabel} di bulan ${monthLabel} ${selectedYear}?\n\nTindakan ini tidak dapat dibatalkan!`
      )
    ) {
      setIsDeleting(true);
      fetch(endpoint, {
        method: "POST",
        mode: "no-cors",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          type: "deleteAttendanceByFilter",
          kelas: selectedKelas === "Semua" ? "" : selectedKelas,
          bulan: selectedMonth,
          tahun: selectedYear,
        }),
      })
        .then(() => {
          alert(
            `✅ Data absensi ${kelasLabel} di bulan ${monthLabel} ${selectedYear} berhasil dihapus.`
          );
          setAttendanceData([]);
          setEditedRecords({});
          fetchAttendanceData();
        })
        .catch(() =>
          alert(
            `❌ Gagal menghapus data absensi ${kelasLabel} di bulan ${monthLabel} ${selectedYear}.`
          )
        )
        .finally(() => setIsDeleting(false));
    }
  };

  const downloadPDF = async () => {
    const doc = new jsPDF({
      orientation: "landscape",
      unit: "mm",
      format: "legal", // Legal: 215.9 x 355.6 mm
    });
    const pageWidth = doc.internal.pageSize.getWidth();
    const margin = 14;
    const lineSpacing = 5;
    let currentY = margin;

    doc.setFont("Times", "roman");

    // Title - 1 baris saja
    const monthLabel =
      months.find((m) => m.value === selectedMonth)?.label || "";
    const namaSekolah = schoolData?.namaSekolah || "UPT SDN 13 BATANG";

    // Hitung posisi tengah tabel (bukan tengah kertas)
    const tableStartX = margin - 7; // Sesuaikan dengan margin tabel Anda
    const mainTableWidth = pageWidth - 2 * margin; // Lebar total tabel utama
    const tableCenterX = tableStartX + mainTableWidth / 2;

    // Judul dalam 1 baris
    const title = `DAFTAR HADIR SISWA KELAS ${selectedKelas}  ${namaSekolah}  ${monthLabel.toUpperCase()} ${selectedYear}`;

    doc.setFontSize(12); // Ukuran font lebih kecil agar muat 1 baris
    doc.setFont("Times", "bold");
    doc.text(title, tableCenterX, currentY, { align: "center" });

    currentY += 10; // Spacing setelah judul (1 baris saja)

    // TAMBAHKAN helper function untuk cek rentang tanggal di dalam downloadPDF
    const isInDateRangePDF = (
      day: number,
      startDate: string,
      endDate?: string
    ): boolean => {
      const currentDate = `${String(day).padStart(2, "0")}/${String(
        selectedMonth
      ).padStart(2, "0")}/${selectedYear}`;

      if (!endDate) {
        return currentDate === startDate;
      }

      // Parse dates
      const [d1, m1, y1] = startDate.split("/").map(Number);
      const [d2, m2, y2] = endDate.split("/").map(Number);
      const [dC, mC, yC] = currentDate.split("/").map(Number);

      const start = new Date(y1, m1 - 1, d1);
      const end = new Date(y2, m2 - 1, d2);
      const current = new Date(yC, mC - 1, dC);

      return current >= start && current <= end;
    };

    // TAMBAHKAN helper untuk cek libur semester di PDF
    const isLiburSemesterPDF = (day: number): boolean => {
      return tanggalMerahList.some((tm) => {
        const isLibur =
          tm.deskripsi.toLowerCase().includes("libur akhir semester") ||
          tm.deskripsi.toLowerCase().includes("libur semester");

        if (!isLibur) return false;

        return isInDateRangePDF(day, tm.tanggal, tm.tanggalAkhir);
      });
    };

    // TAMBAHKAN helper untuk cek tanggal merah biasa di PDF
    const isTanggalMerahPDF = (day: number): boolean => {
      return tanggalMerahList.some((tm) => {
        const isNotLiburSemester = !(
          tm.deskripsi.toLowerCase().includes("libur akhir semester") ||
          tm.deskripsi.toLowerCase().includes("libur semester")
        );

        if (!isNotLiburSemester) return false;

        return isInDateRangePDF(day, tm.tanggal, tm.tanggalAkhir);
      });
    };

    // Headers
    const headers = [
      [
        { content: "No", rowSpan: 2 },
        { content: "NISN", rowSpan: 2 },
        { content: "NAMA", rowSpan: 2 },
        { content: "L/P", rowSpan: 2 },
        ...Array.from({ length: daysInMonth }, (_, i) => ({
          content: (i + 1).toString(),
          rowSpan: 2,
        })),
        {
          content: "JUMLAH",
          colSpan: 4,
          styles: { halign: "center" as const },
        },
      ],
      ["H", "S", "I", "A"],
    ];

    // Body data untuk siswa
    const body = filteredStudents.map((student, index) => {
      const cached = studentAttendanceMap.get(student.id);
      if (!cached) return [];

      const { attendance, counts } = cached;

      // ========== TAMBAHKAN KODE BARU INI ==========
      const jenisKelamin =
        student.jenisKelamin === "L" || student.jenisKelamin === "LAKI-LAKI"
          ? "L"
          : student.jenisKelamin === "P" || student.jenisKelamin === "PEREMPUAN"
          ? "P"
          : "-";
      // ========== AKHIR KODE BARU ==========

      return [
        index + 1,
        student.nisn || "N/A",
        student.name || "N/A",
        jenisKelamin,
        ...Array.from(
          { length: daysInMonth },
          (_, day) => attendance[day + 1] || "-"
        ),
        counts.H,
        counts.S,
        counts.I,
        counts.A,
      ];
    });

    // Setelah baris percentRow, tambahkan:
    const totalSummary = getTotalSummary();

    const totalRow = [
      {
        content: "TOTAL",
        colSpan: 4 + daysInMonth, // ← UBAH: Gabungkan dari No, NISN, Nama, L/P sampai tanggal terakhir
        styles: { halign: "center" as const, fontStyle: "bold" as const },
      },
      totalSummary.totalH,
      totalSummary.totalS,
      totalSummary.totalI,
      totalSummary.totalA,
    ];

    const percentTotalRow = [
      {
        content: "PERSENTASE BULANAN",
        colSpan: 4 + daysInMonth, // ← UBAH: Gabungkan dari No, NISN, Nama, L/P sampai tanggal terakhir
        styles: { halign: "center" as const, fontStyle: "bold" as const },
      },
      `${totalSummary.percentH}%`,
      `${totalSummary.percentS}%`,
      `${totalSummary.percentI}%`,
      `${totalSummary.percentA}%`,
    ];

    // Hitung jumlah hadir per tanggal
    const attendanceByDate = getAttendanceByDate();

    // Baris Jumlah Hadir
    const jumlahHadirRow = [
      {
        content: "Jumlah Hadir",
        colSpan: 4,
        styles: { halign: "center" as const, fontStyle: "bold" as const },
      },
      ...Array.from({ length: daysInMonth }, (_, day) => {
        const dayNum = day + 1;
        const stats = attendanceByDate[dayNum] || { hadir: 0, total: 0 };
        return stats.total > 0 ? stats.hadir : "";
      }),
      "-",
      "-",
      "-",
      "-",
    ];

    const persenHadirRow = [
      {
        content: "% Hadir",
        colSpan: 4,
        styles: { halign: "center" as const, fontStyle: "bold" as const },
      },
      ...Array.from({ length: daysInMonth }, (_, day) => {
        const dayNum = day + 1;
        const stats = attendanceByDate[dayNum] || { hadir: 0, total: 0 };
        const percentage =
          stats.total > 0
            ? ((stats.hadir / stats.total) * 100).toFixed(0) + "%"
            : "";
        return percentage;
      }),
      "-",
      "-",
      "-",
      "-",
    ];

    // Hitung hari efektif untuk PDF (sama seperti di tabel)
    let hariEfektifPDF = 0;
    for (let day = 1; day <= daysInMonth; day++) {
      const isSundayDay = isSunday(day);
      const isTglMerah = isTanggalMerahPDF(day);
      const isLiburSem = isLiburSemesterPDF(day);
      const isBukanJadwal = isBukanJadwalMengajar(day); // ✅ TAMBAHKAN

      // ✅ TAMBAHKAN && !isBukanJadwal
      if (!isSundayDay && !isTglMerah && !isLiburSem && !isBukanJadwal) {
        hariEfektifPDF++;
      }
    }

    const hariEfektifRow = [
      {
        content: "HARI EFEKTIF",
        colSpan: 4 + daysInMonth,
        styles: { halign: "center" as const, fontStyle: "bold" as const },
      },
      {
        content: `${hariEfektifPDF} Hari`,
        colSpan: 4,
        styles: { halign: "center" as const, fontStyle: "bold" as const },
      },
    ];

    // Render tabel dengan baris tambahan
    autoTable(doc, {
      head: headers,
      body: [
        ...body,
        jumlahHadirRow,
        persenHadirRow,
        totalRow,
        percentTotalRow,
        hariEfektifRow,
      ],
      startY: currentY,
      showHead: "firstPage",
      margin: { left: 8, right: 8 },
      theme: "grid",

      styles: {
        font: "Times",
        fontSize: 7,
        cellPadding: 1.5,
        halign: "center",
        valign: "middle",
      },
      headStyles: {
        fillColor: [255, 255, 0],
        textColor: [0, 0, 0],
        fontStyle: "bold",
        fontSize: 7,
        halign: "center",
        valign: "middle",
      },
      alternateRowStyles: { fillColor: [240, 240, 240] },

      columnStyles: {
        0: { cellWidth: 7, halign: "center" }, // No
        1: { cellWidth: 16, halign: "center" }, // NISN
        2: { cellWidth: 60, halign: "left" }, // Nama
        3: { cellWidth: 8, halign: "center" },
        ...Object.assign(
          {},
          ...Array.from({ length: daysInMonth }, (_, i) => ({
            [i + 4]: { cellWidth: 6, halign: "center" },
          }))
        ),
        [4 + daysInMonth]: { cellWidth: 10, halign: "center" },
        [5 + daysInMonth]: { cellWidth: 10, halign: "center" },
        [6 + daysInMonth]: { cellWidth: 10, halign: "center" },
        [7 + daysInMonth]: { cellWidth: 10, halign: "center" },
      },
      didParseCell: function (data) {
        // Styling khusus untuk baris "Jumlah Hadir" dan "% Hadir"
        if (data.row.index >= body.length) {
          data.cell.styles.fillColor =
            data.row.index === body.length ? [219, 234, 254] : [220, 252, 231];
          data.cell.styles.fontStyle = "bold";
        }

        // Styling untuk kolom tanggal di HEADER
        if (
          data.row.section === "head" &&
          data.column.index >= 4 &&
          data.column.index < 4 + daysInMonth
        ) {
          const dayNum = data.column.index - 3;

          const date = new Date(selectedYear, selectedMonth - 1, dayNum);
          const isSundayDay = date.getDay() === 0;

          // UBAH: Gunakan helper function yang baru
          const isTglMerah = isTanggalMerahPDF(dayNum);
          const isLiburSem = isLiburSemesterPDF(dayNum);
          const isBukanJadwal = isBukanJadwalMengajar(dayNum);

          // TAMBAHKAN helper function di awal downloadPDF
          const hexToRgb = (hex: string): [number, number, number] => {
            const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(
              hex
            );
            return result
              ? [
                  parseInt(result[1], 16),
                  parseInt(result[2], 16),
                  parseInt(result[3], 16),
                ]
              : [255, 255, 255];
          };

          // Kemudian gunakan di didParseCell:
          if (isSundayDay) {
            data.cell.styles.fillColor = hexToRgb(customColors.hariMinggu);
            data.cell.styles.textColor = [255, 255, 255];
            data.cell.styles.fontStyle = "bold";
          } else if (isLiburSem) {
            data.cell.styles.fillColor = hexToRgb(customColors.liburSemester);
            data.cell.styles.textColor = [255, 255, 255];
            data.cell.styles.fontStyle = "bold";
          } else if (isTglMerah) {
            data.cell.styles.fillColor = hexToRgb(customColors.tanggalMerah);
            data.cell.styles.textColor = [0, 0, 0];
            data.cell.styles.fontStyle = "bold";
          } else if (isBukanJadwal) {
            data.cell.styles.fillColor = hexToRgb(customColors.bukanJadwal);
            data.cell.styles.textColor = [0, 0, 0];
            data.cell.styles.fontStyle = "bold";
          }
        }

        // Styling untuk cell data siswa di BODY
        if (
          data.row.section === "body" &&
          data.row.index < body.length &&
          data.column.index >= 4 &&
          data.column.index < 4 + daysInMonth
        ) {
          const dayNum = data.column.index - 3;

          const date = new Date(selectedYear, selectedMonth - 1, dayNum);
          const isSundayDay = date.getDay() === 0;

          // UBAH: Gunakan helper function yang baru
          const isTglMerah = isTanggalMerahPDF(dayNum);
          const isLiburSem = isLiburSemesterPDF(dayNum);
          const isBukanJadwal = isBukanJadwalMengajar(dayNum);

          // TAMBAHKAN helper function di awal downloadPDF
          const hexToRgb = (hex: string): [number, number, number] => {
            const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(
              hex
            );
            return result
              ? [
                  parseInt(result[1], 16),
                  parseInt(result[2], 16),
                  parseInt(result[3], 16),
                ]
              : [255, 255, 255];
          };

          // Kemudian gunakan di didParseCell:
          if (isSundayDay) {
            data.cell.styles.fillColor = hexToRgb(customColors.hariMinggu);
            data.cell.styles.textColor = [255, 255, 255];
            data.cell.styles.fontStyle = "bold";
          } else if (isLiburSem) {
            data.cell.styles.fillColor = hexToRgb(customColors.liburSemester);
            data.cell.styles.textColor = [255, 255, 255];
            data.cell.styles.fontStyle = "bold";
          } else if (isTglMerah) {
            data.cell.styles.fillColor = hexToRgb(customColors.tanggalMerah);
            data.cell.styles.textColor = [0, 0, 0];
            data.cell.styles.fontStyle = "bold";
          } else if (isBukanJadwal) {
            data.cell.styles.fillColor = hexToRgb(customColors.bukanJadwal);
            data.cell.styles.textColor = [0, 0, 0];
            data.cell.styles.fontStyle = "bold";
          }
        }

        // Styling untuk footer rows pada tanggal merah
        if (
          data.row.section === "body" &&
          data.row.index >= body.length &&
          data.column.index >= 4 &&
          data.column.index < 4 + daysInMonth
        ) {
          const dayNum = data.column.index - 3;

          const date = new Date(selectedYear, selectedMonth - 1, dayNum);
          const isSundayDay = date.getDay() === 0;

          // UBAH: Gunakan helper function yang baru
          const isTglMerah = isTanggalMerahPDF(dayNum);
          const isLiburSem = isLiburSemesterPDF(dayNum);
          const isBukanJadwal = isBukanJadwalMengajar(dayNum);

          // TAMBAHKAN helper function di awal downloadPDF
          const hexToRgb = (hex: string): [number, number, number] => {
            const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(
              hex
            );
            return result
              ? [
                  parseInt(result[1], 16),
                  parseInt(result[2], 16),
                  parseInt(result[3], 16),
                ]
              : [255, 255, 255];
          };

          // Kemudian gunakan di didParseCell:
          if (isSundayDay) {
            data.cell.styles.fillColor = hexToRgb(customColors.hariMinggu);
            data.cell.styles.textColor = [255, 255, 255];
            data.cell.styles.fontStyle = "bold";
          } else if (isLiburSem) {
            data.cell.styles.fillColor = hexToRgb(customColors.liburSemester);
            data.cell.styles.textColor = [255, 255, 255];
            data.cell.styles.fontStyle = "bold";
          } else if (isTglMerah) {
            data.cell.styles.fillColor = hexToRgb(customColors.tanggalMerah);
            data.cell.styles.textColor = [0, 0, 0];
            data.cell.styles.fontStyle = "bold";
          } else if (isBukanJadwal) {
            data.cell.styles.fillColor = hexToRgb(customColors.bukanJadwal);
            data.cell.styles.textColor = [0, 0, 0];
            data.cell.styles.fontStyle = "bold";
          }
        }
      },
    });

    currentY = (doc as any).lastAutoTable.finalY + 8;

    // ✅ PENGECEKAN LEBIH KETAT - Cek apakah footer (tanda tangan) akan terpotong
    const pageHeight = doc.internal.pageSize.getHeight();
    const bottomMargin = 20;
    const spaceNeededForSignatures = 50; // Ruang untuk tanda tangan
    const spaceNeededForStudentTable = 20; // Ruang untuk tabel jumlah siswa
    const spaceNeededForTanggalMerahTable =
      filteredTanggalMerah.length > 0
        ? 10 + filteredTanggalMerah.length * 7
        : 0;

    // Total ruang yang dibutuhkan untuk semua footer
    const totalFooterSpace =
      spaceNeededForSignatures +
      spaceNeededForStudentTable +
      spaceNeededForTanggalMerahTable;

    // ✅ JIKA TIDAK CUKUP RUANG UNTUK SEMUA FOOTER, PINDAH KE HALAMAN BARU
    if (currentY + totalFooterSpace > pageHeight - bottomMargin) {
      doc.addPage();
      currentY = margin;
    }

    // TABEL JUMLAH SISWA (sekarang dijamin di halaman yang sama dengan tanda tangan)
    const genderSummary = getGenderSummary();

    doc.setFontSize(10);
    doc.setFont("Times", "bold");
    doc.text("JUMLAH SISWA", margin - 7, currentY, {
      align: "left",
    });
    currentY += 3;

    const tableWidth = (pageWidth - 2 * margin) * 0.4;

    autoTable(doc, {
      head: [["LAKI-LAKI", "PEREMPUAN", "TOTAL SISWA"]],
      body: [
        [
          genderSummary.lakiLaki.toString(),
          genderSummary.perempuan.toString(),
          genderSummary.total.toString(),
        ],
      ],
      startY: currentY,
      margin: { left: margin - 7, right: pageWidth - margin - tableWidth },
      tableWidth: tableWidth,
      theme: "grid",
      styles: {
        font: "Times",
        fontSize: 7,
        cellPadding: 1,
        halign: "center",
        valign: "middle",
        lineWidth: 0.5,
      },
      headStyles: {
        fillColor: [255, 255, 255],
        textColor: [0, 0, 0],
        fontStyle: "bold",
        lineWidth: 1,
      },
      bodyStyles: {
        fillColor: [255, 255, 255],
        textColor: [0, 0, 0],
        fontStyle: "bold",
        fontSize: 10,
        lineWidth: 1,
      },
      columnStyles: {
        0: {
          cellWidth: tableWidth / 3,
          fillColor: [255, 255, 255],
        },
        1: {
          cellWidth: tableWidth / 3,
          fillColor: [255, 255, 255],
        },
        2: {
          cellWidth: tableWidth / 3,
          fillColor: [255, 255, 255],
        },
      },
    });

    currentY = (doc as any).lastAutoTable.finalY + 10;

    // TABEL KETERANGAN TANGGAL MERAH (masih di halaman yang sama)
    if (filteredTanggalMerah.length > 0) {
      doc.setFontSize(10);
      doc.setFont("Times", "bold");
      doc.text("KETERANGAN TANGGAL MERAH / LIBUR", margin - 7, currentY, {
        align: "left",
      });
      currentY += 3;

      const sortedTanggalMerah = [...filteredTanggalMerah].sort((a, b) => {
        const [dayA, monthA, yearA] = a.tanggal.split("/").map(Number);
        const [dayB, monthB, yearB] = b.tanggal.split("/").map(Number);
        return (
          new Date(yearA, monthA - 1, dayA).getTime() -
          new Date(yearB, monthB - 1, dayB).getTime()
        );
      });

      autoTable(doc, {
        head: [["TANGGAL", "KETERANGAN"]],
        body: sortedTanggalMerah.map((tm) => [
          tm.tanggalAkhir ? `${tm.tanggal} - ${tm.tanggalAkhir}` : tm.tanggal,
          tm.deskripsi,
        ]),
        startY: currentY,
        margin: { left: margin - 7, right: margin - 7 },
        theme: "grid",
        styles: {
          font: "Times",
          fontSize: 8,
          cellPadding: 1.5,
          valign: "middle",
          lineWidth: 0.5,
        },
        headStyles: {
          fillColor: [254, 202, 202],
          textColor: [0, 0, 0],
          fontStyle: "bold",
          halign: "center",
          fontSize: 9,
          lineWidth: 1,
        },
        bodyStyles: {
          fillColor: [255, 255, 255],
          textColor: [0, 0, 0],
          lineWidth: 1,
        },
        alternateRowStyles: {
          fillColor: [254, 226, 226],
        },
        columnStyles: {
          0: {
            cellWidth: 35,
            halign: "center",
            fontSize: 8,
          },
          1: {
            cellWidth: 96,
            halign: "left",
            fontSize: 8,
          },
        },
        tableWidth: "auto",
        didDrawCell: function (data) {
          if (data.column.index === 1 && data.cell.section === "body") {
            data.cell.styles.cellWidth = 120;
          }
        },
      });

      currentY = (doc as any).lastAutoTable.finalY + 10;
    }

    // Footer: School data, place, date, signatures
    if (schoolData) {
      doc.setFontSize(10);
      doc.setFont("Times", "roman");

      const formattedDate = new Date(selectedDate).toLocaleDateString("id-ID", {
        day: "2-digit",
        month: "long",
        year: "numeric",
      });
      const placeDateText = `${
        schoolData.namaKota || "Makassar"
      }, ${formattedDate}`;
      const rightColumnX = pageWidth / 2 + 80; // 👈 DIUBAH - posisi guru digeser ke kiri
      doc.text(placeDateText, rightColumnX + 25, currentY - 1, {
        align: "center",
      });
      currentY += 5;

      const signatureWidth = 30;
      const signatureHeight = 20;
      const leftColumnX = margin;

      if (schoolData.ttdKepsek) {
        try {
          const canvas = document.createElement("canvas");
          canvas.width = 150; // Sesuaikan ukuran canvas (lebar lebih besar untuk tanda tangan panjang)
          canvas.height = 50; // Sesuaikan ukuran canvas (tinggi cukup untuk garis tanda tangan)
          const ctx = canvas.getContext("2d");
          const v = await Canvg.from(ctx, schoolData.ttdKepsek); // schoolData.ttdKepsek adalah base64 SVG
          v.start();
          const pngData = canvas.toDataURL("image/png");
          doc.addImage(
            pngData,
            "PNG",
            leftColumnX + 10,
            currentY - 3,
            signatureWidth,
            signatureHeight
          ); // Sesuaikan posisi sesuai asli
        } catch (error) {
          console.error("Error rendering Kepsek signature:", error);
          doc.setFontSize(10);
          doc.text(
            "Gagal render tanda tangan Kepala Sekolah.",
            leftColumnX + 10,
            currentY - 3 + 10
          );
        }
      }

      doc.text("Kepala Sekolah,", leftColumnX + 25, currentY - 2, {
        align: "center",
      });
      doc.text("", leftColumnX + 25, currentY + lineSpacing, {
        align: "center",
      });
      doc.text("", leftColumnX + 25, currentY + 2 * lineSpacing, {
        align: "center",
      });

      const principalName = schoolData.namaKepsek || "N/A";
      doc.setFont("Times", "bold");
      doc.text(principalName, leftColumnX + 25, currentY + 3.5 * lineSpacing, {
        align: "center",
      });

      const textWidth = doc.getTextWidth(principalName);
      const textX = leftColumnX + 25 - textWidth / 2;
      doc.line(
        textX,
        currentY + 3.5 * lineSpacing + 1,
        textX + textWidth,
        currentY + 3.5 * lineSpacing + 1
      );

      doc.setFont("Times", "roman");
      doc.text(
        `NIP. ${schoolData.nipKepsek || "N/A"}`,
        leftColumnX + 25,
        currentY + 4.5 * lineSpacing,
        { align: "center" }
      );

      if (schoolData.ttdGuru) {
        try {
          const canvas = document.createElement("canvas");
          canvas.width = 150; // Sesuaikan ukuran canvas
          canvas.height = 50;
          const ctx = canvas.getContext("2d");
          const v = await Canvg.from(ctx, schoolData.ttdGuru); // schoolData.ttdGuru adalah base64 SVG
          v.start();
          const pngData = canvas.toDataURL("image/png");
          doc.addImage(
            pngData,
            "PNG",
            rightColumnX + 10,
            currentY - 5,
            signatureWidth,
            signatureHeight
          ); // Sesuaikan posisi sesuai asli
        } catch (error) {
          console.error("Error rendering Guru signature:", error);
          doc.setFontSize(10);
          doc.text(
            "Gagal render tanda tangan Guru.",
            rightColumnX + 10,
            currentY - 5 + 10
          );
        }
      }

      doc.text(
        `${schoolData.statusGuru || "Guru Kelas"},`,
        rightColumnX + 25,
        currentY - 2,
        {
          align: "center",
        }
      );
      doc.text("", rightColumnX + 25, currentY + lineSpacing, {
        align: "center",
      });
      doc.text("", rightColumnX + 25, currentY + 2 * lineSpacing, {
        align: "center",
      });

      const teacherName = schoolData.namaGuru || "N/A";
      doc.setFont("Times", "bold");
      doc.text(teacherName, rightColumnX + 25, currentY + 3.5 * lineSpacing, {
        align: "center",
      });

      const teacherTextWidth = doc.getTextWidth(teacherName);
      const teacherTextX = rightColumnX + 25 - teacherTextWidth / 2;
      doc.line(
        teacherTextX,
        currentY + 3.5 * lineSpacing + 1,
        teacherTextX + teacherTextWidth,
        currentY + 3.5 * lineSpacing + 1
      );

      doc.setFont("Times", "roman");
      doc.text(
        `NIP. ${schoolData.nipGuru || "N/A"}`,
        rightColumnX + 25,
        currentY + 4.5 * lineSpacing,
        { align: "center" }
      );
    } else {
      doc.setFontSize(10);
      doc.text("Data sekolah tidak tersedia.", margin, currentY);
    }

    const date = new Date()
      .toLocaleString("id-ID", {
        day: "2-digit",
        month: "long",
        year: "numeric",
        hour: "2-digit",
        minute: "2-digit",
        hour12: false,
      })
      .replace(/ /g, "_")
      .replace(/:/g, "-");
    const fileName = `Daftar_Hadir_${selectedKelas}_${monthLabel}_${selectedYear}_${date}.pdf`;
    doc.save(fileName);
  };

  const handleNameClick = (student: Student) => {
    setSelectedStudent(student);
    setShowModal(true);
  };

  const closeModal = () => {
    setShowModal(false);
    setSelectedStudent(null);
  };

  if (loading) {
    return (
      <div className="text-center py-8">
        <p className="text-gray-500">Memuat data...</p>
      </div>
    );
  }

  return (
    <div className="max-w-4xl mx-auto" style={{ paddingBottom: "70px" }}>
      <div className="bg-white p-6 rounded-lg shadow-md">
        <h2 className="text-2xl font-bold text-center text-blue-700 mb-6">
          📋 Daftar Hadir Siswa
        </h2>
        <div className="mb-6 flex flex-col md:flex-row gap-4 items-center justify-center">
          <div className="text-center">
            <p className="text-sm text-gray-500 mb-2">Filter Kelas</p>
            <select
              value={selectedKelas}
              onChange={(e) => setSelectedKelas(e.target.value)}
              className="border border-gray-300 rounded-lg px-1 py-0.5 shadow-sm bg-white min-w-32"
            >
              {uniqueClasses.map((kelas) => (
                <option key={kelas} value={kelas}>
                  {kelas}
                </option>
              ))}
            </select>
          </div>
          <div className="text-center">
            <p className="text-sm text-gray-500 mb-2">Bulan</p>
            <select
              value={selectedMonth}
              onChange={(e) => setSelectedMonth(Number(e.target.value))}
              className="border border-gray-300 rounded-lg px-1 py-0.5 shadow-sm bg-white min-w-32"
            >
              {months.map((m) => (
                <option key={m.value} value={m.value}>
                  {m.label}
                </option>
              ))}
            </select>
          </div>
          <div className="text-center">
            <p className="text-sm text-gray-500 mb-2">Tahun</p>
            <select
              value={selectedYear}
              onChange={(e) => setSelectedYear(Number(e.target.value))}
              className="border border-gray-300 rounded-lg px-1 py-0.5 shadow-sm bg-white min-w-32"
            >
              {years.map((y) => (
                <option key={y} value={y}>
                  {y}
                </option>
              ))}
            </select>
          </div>
        </div>

        {/* Separator line and PDF settings section */}
        <div className="border-t border-gray-200 pt-4 mb-6">
          <p className="text-center text-sm font-medium text-gray-700 mb-4">
            Pengaturan Tanggal & Nama Tempat <br /> untuk Daftar Hadir pada File
            PDF
          </p>
          <div className="flex flex-col md:flex-row gap-4 items-center justify-center">
            <div className="text-center">
              <p className="text-sm text-gray-500 mb-2">Pilih Tanggal</p>
              <input
                type="date"
                value={selectedDate}
                onChange={(e) => setSelectedDate(e.target.value)}
                className="border border-gray-300 rounded-lg px-1 py-0.5 shadow-sm bg-white min-w-32"
              />
            </div>
          </div>
        </div>

        {/* ========== SECTION BARU: PENGATURAN WARNA ========== */}
        <div className="border-t border-gray-200 pt-4 mb-6">
          <div className="flex items-center justify-between mb-4">
            <p className="text-sm font-medium text-gray-700">
              🎨 Pengaturan Warna Tabel
            </p>
            <button
              onClick={() => setShowColorPicker(!showColorPicker)}
              className="text-sm text-blue-600 hover:text-blue-800 font-medium"
            >
              {showColorPicker ? "▲ Sembunyikan" : "▼ Tampilkan"}
            </button>
          </div>

          {showColorPicker && (
            <div className="bg-gray-50 border border-gray-200 rounded-lg p-4">
              <div className="grid grid-cols-2 md:grid-cols-3 gap-4 mb-4">
                {/* Hari Minggu */}
                <div className="flex flex-col gap-2">
                  <label className="text-xs font-medium text-gray-700">
                    Hari Minggu (Disabled)
                  </label>
                  <div className="flex items-center gap-2">
                    <input
                      type="color"
                      value={customColors.hariMinggu}
                      onChange={(e) =>
                        handleColorChange("hariMinggu", e.target.value)
                      }
                      className="w-12 h-10 border border-gray-300 rounded cursor-pointer"
                    />
                    <span className="text-xs text-gray-600">
                      {customColors.hariMinggu}
                    </span>
                  </div>
                </div>

                {/* Libur Semester */}
                <div className="flex flex-col gap-2">
                  <label className="text-xs font-medium text-gray-700">
                    Libur Semester (Disabled)
                  </label>
                  <div className="flex items-center gap-2">
                    <input
                      type="color"
                      value={customColors.liburSemester}
                      onChange={(e) =>
                        handleColorChange("liburSemester", e.target.value)
                      }
                      className="w-12 h-10 border border-gray-300 rounded cursor-pointer"
                    />
                    <span className="text-xs text-gray-600">
                      {customColors.liburSemester}
                    </span>
                  </div>
                </div>

                {/* Tanggal Merah */}
                <div className="flex flex-col gap-2">
                  <label className="text-xs font-medium text-gray-700">
                    Tanggal Merah (Libur)
                  </label>
                  <div className="flex items-center gap-2">
                    <input
                      type="color"
                      value={customColors.tanggalMerah}
                      onChange={(e) =>
                        handleColorChange("tanggalMerah", e.target.value)
                      }
                      className="w-12 h-10 border border-gray-300 rounded cursor-pointer"
                    />
                    <span className="text-xs text-gray-600">
                      {customColors.tanggalMerah}
                    </span>
                  </div>
                </div>

                {/* Bukan Jadwal */}
                {schoolData?.statusGuru !== "Guru Kelas" &&
                  selectedKelas !== "Semua" && (
                    <div className="flex flex-col gap-2">
                      <label className="text-xs font-medium text-gray-700">
                        Bukan Jadwal Mengajar
                      </label>
                      <div className="flex items-center gap-2">
                        <input
                          type="color"
                          value={customColors.bukanJadwal}
                          onChange={(e) =>
                            handleColorChange("bukanJadwal", e.target.value)
                          }
                          className="w-12 h-10 border border-gray-300 rounded cursor-pointer"
                        />
                        <span className="text-xs text-gray-600">
                          {customColors.bukanJadwal}
                        </span>
                      </div>
                    </div>
                  )}

                {/* Jadwal Mengajar */}
                <div className="flex flex-col gap-2">
                  <label className="text-xs font-medium text-gray-700">
                    {schoolData?.statusGuru === "Guru Kelas" ||
                    selectedKelas === "Semua"
                      ? "Hari Normal"
                      : "Jadwal Mengajar"}
                  </label>
                  <div className="flex items-center gap-2">
                    <input
                      type="color"
                      value={customColors.jadwalMengajar}
                      onChange={(e) =>
                        handleColorChange("jadwalMengajar", e.target.value)
                      }
                      className="w-12 h-10 border border-gray-300 rounded cursor-pointer"
                    />
                    <span className="text-xs text-gray-600">
                      {customColors.jadwalMengajar}
                    </span>
                  </div>
                </div>

                {/* Data Edit */}
                <div className="flex flex-col gap-2">
                  <label className="text-xs font-medium text-gray-700">
                    Data Sedang Diedit
                  </label>
                  <div className="flex items-center gap-2">
                    <input
                      type="color"
                      value={customColors.dataEdit}
                      onChange={(e) =>
                        handleColorChange("dataEdit", e.target.value)
                      }
                      className="w-12 h-10 border border-gray-300 rounded cursor-pointer"
                    />
                    <span className="text-xs text-gray-600">
                      {customColors.dataEdit}
                    </span>
                  </div>
                </div>
              </div>

              {/* Button Reset */}
              <div className="flex justify-end">
                <button
                  onClick={resetColors}
                  className="text-xs px-4 py-2 bg-gray-500 hover:bg-gray-600 text-white rounded-lg font-medium transition-colors"
                >
                  🔄 Reset ke Warna Default
                </button>
              </div>
            </div>
          )}
        </div>
        {/* ========== AKHIR SECTION PENGATURAN WARNA ========== */}

        {/* Legenda Warna - UPDATE DENGAN WARNA DINAMIS */}
        <div className="mb-4 bg-gray-50 border border-gray-300 rounded-lg p-4">
          <h4 className="text-sm font-semibold text-gray-700 mb-3">
            📌 Keterangan Warna:
          </h4>
          <div className="grid grid-cols-2 md:grid-cols-3 gap-3">
            <div className="flex items-center gap-2">
              <div
                className="w-6 h-6 border border-gray-300 rounded"
                style={{ backgroundColor: customColors.hariMinggu }}
              ></div>
              <span className="text-xs text-gray-700">
                Hari Minggu (Disabled)
              </span>
            </div>
            <div className="flex items-center gap-2">
              <div
                className="w-6 h-6 border border-gray-300 rounded"
                style={{ backgroundColor: customColors.liburSemester }}
              ></div>
              <span className="text-xs text-gray-700">
                Libur Semester (Disabled)
              </span>
            </div>
            <div className="flex items-center gap-2">
              <div
                className="w-6 h-6 border border-gray-300 rounded"
                style={{ backgroundColor: customColors.tanggalMerah }}
              ></div>
              <span className="text-xs text-gray-700">
                Tanggal Merah (Libur)
              </span>
            </div>
            {schoolData?.statusGuru !== "Guru Kelas" &&
              selectedKelas !== "Semua" && (
                <div className="flex items-center gap-2">
                  <div
                    className="w-6 h-6 border border-gray-300 rounded"
                    style={{ backgroundColor: customColors.bukanJadwal }}
                  ></div>
                  <span className="text-xs text-gray-700">
                    Bukan Jadwal Mengajar
                  </span>
                </div>
              )}
            <div className="flex items-center gap-2">
              <div
                className="w-6 h-6 border border-gray-300 rounded"
                style={{ backgroundColor: customColors.jadwalMengajar }}
              ></div>
              <span className="text-xs text-gray-700">
                {schoolData?.statusGuru === "Guru Kelas" ||
                selectedKelas === "Semua"
                  ? "Hari Normal"
                  : "Jadwal Mengajar"}
              </span>
            </div>
            <div className="flex items-center gap-2">
              <div
                className="w-6 h-6 border border-gray-300 rounded"
                style={{ backgroundColor: customColors.dataEdit }}
              ></div>
              <span className="text-xs text-gray-700">Data Sedang Diedit</span>
            </div>
          </div>
        </div>

        <div className="overflow-auto relative" style={{ maxHeight: "70vh" }}>
          <style>
            {`
              .attendance-table thead tr:first-child th {
    position: sticky;
    top: 0;
    z-index: 30;
    background: #f3f4f6 !important;
  }
  
  .attendance-table thead tr:last-child th {
    position: sticky;
    top: 33px;
    z-index: 30;
    background: #f3f4f6 !important;
  
  .attendance-table th.freeze-no,
  .attendance-table td.freeze-no {
    position: sticky;
    left: 0;
    z-index: 25;
    background: #f3f4f6 !important;
    box-shadow: 2px 0 5px rgba(0,0,0,0.1);
  }
  
  .attendance-table td.freeze-no {
    background: white !important;
  }
  
  .attendance-table tbody tr:nth-child(even) td.freeze-no {
    background: #f9fafb !important;
  }
  
  .attendance-table th.freeze-nama,
  .attendance-table td.freeze-nama {
    position: sticky;
    left: 30px;
    z-index: 25;
    background: #f3f4f6 !important;
    box-shadow: 2px 0 5px rgba(0,0,0,0.1);
    cursor: pointer;
    min-width: 100px;
  }
  
  .attendance-table td.freeze-nama {
    background: white !important;
  }
  
  .attendance-table td.freeze-nama:hover {
    background: #dbeafe !important;
  }
  
  .attendance-table tbody tr:nth-child(even) td.freeze-nama {
    background: #f9fafb !important;
  }
  
  .attendance-table tbody tr:nth-child(even) td.freeze-nama:hover {
    background: #dbeafe !important;
  }

  .attendance-table th.freeze-jk,
  .attendance-table td.freeze-jk {
    position: sticky;
    left: 130px; /* Sesuaikan dengan lebar kolom No + Nama */
    z-index: 25;
    background: #f3f4f6 !important;
    box-shadow: 2px 0 5px rgba(0,0,0,0.1);
    min-width: 40px;
    max-width: 40px;
  }
  
  .attendance-table td.freeze-jk {
    background: white !important;
  }
  
  .attendance-table tbody tr:nth-child(even) td.freeze-jk {
    background: #f9fafb !important;
  }
  
  .attendance-table thead tr:first-child th.freeze-no,
  .attendance-table thead tr:last-child th.freeze-no {
    z-index: 40;
    background: #f3f4f6 !important;
  }
  
  .attendance-table thead tr:first-child th.freeze-nama,
  .attendance-table thead tr:last-child th.freeze-nama {
    z-index: 40;
    background: #f3f4f6 !important;
  }

  .attendance-table thead tr:first-child th.freeze-jk,
  .attendance-table thead tr:last-child th.freeze-jk {
    z-index: 40;
    background: #f3f4f6 !important;
  }
  .attendance-table tfoot td.freeze-no {
    position: sticky;
    left: 0;
    z-index: 25;
    box-shadow: 2px 0 5px rgba(0,0,0,0.1);
  }
  .attendance-table tfoot td.freeze-jk {
    position: sticky;
    left: 130px;
    z-index: 25;
    box-shadow: 2px 0 5px rgba(0,0,0,0.1);
  }
`}
          </style>

          <table className="attendance-table min-w-full border-collapse border border-gray-200">
            <thead>
              <tr className="bg-gray-100">
                <th className="freeze-no border px-2 py-1 text-sm">No</th>
                <th className="freeze-nama border px-2 py-1 text-sm">NAMA</th>
                <th className="freeze-jk border px-2 py-1 text-sm text-center">
                  L/P
                </th>
                {Array.from({ length: daysInMonth }, (_, i) => {
                  const dayNum = i + 1;
                  const hasData = !daysWithNoData.has(dayNum);
                  const isTglMerah = isTanggalMerah(dayNum);
                  const deskripsi = getTanggalMerahDeskripsi(dayNum);
                  const isSundayDay = isSunday(dayNum);
                  const isLiburSem = isLiburSemester(dayNum);
                  const isBukanJadwal = isBukanJadwalMengajar(dayNum); // TAMBAHKAN

                  return (
                    <th
                      key={i}
                      className={`border px-1 py-1 text-sm ${
                        isSundayDay
                          ? "text-black font-bold"
                          : isLiburSem
                          ? "text-black font-bold"
                          : isTglMerah
                          ? "text-black font-bold"
                          : isBukanJadwal
                          ? "text-black font-bold"
                          : ""
                      }`}
                      style={{
                        backgroundColor: isSundayDay
                          ? customColors.hariMinggu
                          : isLiburSem
                          ? customColors.liburSemester
                          : isTglMerah
                          ? customColors.tanggalMerah
                          : isBukanJadwal
                          ? customColors.bukanJadwal
                          : customColors.jadwalMengajar,
                      }}
                      title={
                        isSundayDay
                          ? "Hari Minggu"
                          : isLiburSem
                          ? deskripsi
                          : isTglMerah
                          ? deskripsi
                          : isBukanJadwal // TAMBAHKAN
                          ? "Bukan Jadwal Mengajar"
                          : ""
                      }
                    >
                      {String(dayNum).padStart(2, "0")}
                    </th>
                  );
                })}
                <th className="border px-2 py-1 text-sm" colSpan={4}>
                  JUMLAH
                </th>
              </tr>
              <tr className="bg-gray-100">
                <th className="freeze-no border px-2 py-1 text-sm"></th>
                <th className="freeze-nama border px-2 py-1 text-sm"></th>
                <th className="freeze-jk border px-2 py-1 text-sm"></th>
                {Array.from({ length: daysInMonth }, (_, i) => {
                  const dayNum = i + 1;
                  const isTglMerah = isTanggalMerah(dayNum);
                  const isSundayDay = isSunday(dayNum);
                  const isLiburSem = isLiburSemester(dayNum);
                  const isBukanJadwal = isBukanJadwalMengajar(dayNum); // TAMBAHKAN

                  return (
                    <th
                      key={i}
                      className={`border px-1 py-1 text-sm ${
                        isSundayDay
                          ? "text-black font-bold"
                          : isLiburSem
                          ? "text-black font-bold"
                          : isTglMerah
                          ? "text-black font-bold"
                          : isBukanJadwal
                          ? "text-black font-bold"
                          : ""
                      }`}
                      style={{
                        backgroundColor: isSundayDay
                          ? customColors.hariMinggu
                          : isLiburSem
                          ? customColors.liburSemester
                          : isTglMerah
                          ? customColors.tanggalMerah
                          : isBukanJadwal
                          ? customColors.bukanJadwal
                          : customColors.jadwalMengajar,
                      }}
                    ></th>
                  );
                })}
                <th className="border px-2 py-1 text-sm">H</th>
                <th className="border px-2 py-1 text-sm">S</th>
                <th className="border px-2 py-1 text-sm">I</th>
                <th className="border px-2 py-1 text-sm">A</th>
              </tr>
            </thead>
            <tbody>
              {filteredStudents.map((student, index) => {
                const cached = studentAttendanceMap.get(student.id);
                if (!cached) return null;

                const { attendance, counts } = cached;

                return (
                  <tr
                    key={student.id}
                    className={index % 2 === 0 ? "bg-white" : "bg-gray-50"}
                  >
                    <td className="freeze-no border px-2 py-1 text-xs text-center">
                      {index + 1}
                    </td>
                    <td
                      className="freeze-nama border px-2 py-1 text-xs"
                      onClick={() => handleNameClick(student)}
                      title="Klik untuk melihat detail"
                    >
                      {student.name || "N/A"}
                    </td>
                    <td className="freeze-jk border px-2 py-1 text-xs text-center font-semibold">
                      {student.jenisKelamin === "L" ||
                      student.jenisKelamin === "LAKI-LAKI"
                        ? "L"
                        : student.jenisKelamin === "P" ||
                          student.jenisKelamin === "PEREMPUAN"
                        ? "P"
                        : "-"}
                    </td>
                    {Array.from({ length: daysInMonth }, (_, day) => {
                      const currentValue = attendance[day + 1] || "";
                      const key = `${student.id}_${day + 1}`;
                      const isEdited = editedRecords[key] !== undefined;
                      const dayNum = day + 1;

                      const hasDataOnThisDate = !daysWithNoData.has(dayNum);
                      const isTglMerah = isTanggalMerah(dayNum);
                      const isSundayDay = isSunday(dayNum);
                      const isLiburSem = isLiburSemester(dayNum);
                      const isBukanJadwal = isBukanJadwalMengajar(dayNum); // TAMBAHKAN

                      const getFullStatus = (
                        code: string
                      ): AttendanceStatus | "" => {
                        switch (code) {
                          case "H":
                            return "Hadir";
                          case "I":
                            return "Izin";
                          case "S":
                            return "Sakit";
                          case "A":
                            return "Alpha";
                          default:
                            return "";
                        }
                      };

                      const getColorClass = (code: string) => {
                        switch (code) {
                          case "H":
                            return "text-green-600 hover:bg-green-50";
                          case "I":
                            return "text-yellow-600 hover:bg-yellow-50";
                          case "S":
                            return "text-blue-600 hover:bg-blue-50";
                          case "A":
                            return "text-red-600 hover:bg-red-50";
                          default:
                            return "text-gray-400 hover:bg-gray-50";
                        }
                      };

                      return (
                        <td
                          key={day}
                          className={`border px-1 py-1 text-sm ${
                            isSundayDay
                              ? "text-black font-bold"
                              : isLiburSem
                              ? "text-black font-bold"
                              : isTglMerah
                              ? "text-black font-bold"
                              : isBukanJadwal
                              ? "text-black font-bold"
                              : ""
                          }`}
                          style={{
                            backgroundColor: isSundayDay
                              ? customColors.hariMinggu
                              : isLiburSem
                              ? customColors.liburSemester
                              : isTglMerah
                              ? customColors.tanggalMerah
                              : isBukanJadwal
                              ? customColors.bukanJadwal
                              : customColors.jadwalMengajar,
                          }}
                        >
                          <select
                            value={
                              editedRecords[key]?.status ||
                              getFullStatus(currentValue)
                            }
                            onChange={(e) => {
                              const newStatus = e.target.value as
                                | AttendanceStatus
                                | "";
                              handleStatusChange(student, day + 1, newStatus);
                            }}
                            className={`w-full text-center text-xs font-bold cursor-pointer appearance-none bg-transparent border-none focus:outline-none focus:ring-2 focus:ring-blue-400 rounded px-1 py-0.5 transition-colors ${getColorClass(
                              editedRecords[key]?.status
                                ? editedRecords[key].status === "Hadir"
                                  ? "H"
                                  : editedRecords[key].status === "Izin"
                                  ? "I"
                                  : editedRecords[key].status === "Sakit"
                                  ? "S"
                                  : editedRecords[key].status === "Alpha"
                                  ? "A"
                                  : ""
                                : currentValue
                            )}`}
                            disabled={
                              isSaving ||
                              isSundayDay ||
                              isTglMerah ||
                              isLiburSemester(day + 1) ||
                              !isJadwalMengajar(day + 1) // TAMBAHKAN - disable jika bukan jadwal
                            }
                            style={{
                              textAlign: "center",
                              textAlignLast: "center",
                              WebkitAppearance: "none",
                              MozAppearance: "none",
                            }}
                          >
                            <option value="">-</option>
                            <option value="Hadir" style={{ color: "#059669" }}>
                              H
                            </option>
                            <option value="Izin" style={{ color: "#D97706" }}>
                              I
                            </option>
                            <option value="Sakit" style={{ color: "#2563EB" }}>
                              S
                            </option>
                            <option value="Alpha" style={{ color: "#DC2626" }}>
                              A
                            </option>
                          </select>
                        </td>
                      );
                    })}
                    <td className="border px-2 py-1 text-center text-xs">
                      {counts.H}
                    </td>
                    <td className="border px-2 py-1 text-center text-xs">
                      {counts.S}
                    </td>
                    <td className="border px-2 py-1 text-center text-xs">
                      {counts.I}
                    </td>
                    <td className="border px-2 py-1 text-center text-xs">
                      {counts.A}
                    </td>
                  </tr>
                );
              })}
            </tbody>
            <tfoot>
              {/* Baris Jumlah Hadir per Tanggal */}
              <tr className="bg-blue-50 font-semibold">
                <td
                  className="freeze-no border px-2 py-1 text-xs text-center"
                  colSpan={3}
                >
                  Jumlah Hadir
                </td>
                {Array.from({ length: daysInMonth }, (_, day) => {
                  const dayNum = day + 1;
                  const stats = getAttendanceByDate()[dayNum] || {
                    hadir: 0,
                    total: 0,
                  };
                  const isTglMerah = isTanggalMerah(dayNum);
                  const isSundayDay = isSunday(dayNum);
                  const isLiburSem = isLiburSemester(dayNum); // TAMBAHAN BARU
                  const isBukanJadwal = isBukanJadwalMengajar(dayNum);

                  return (
                    <td
                      key={day}
                      className={`border px-1 py-1 text-sm ${
                        isSundayDay
                          ? "text-black font-bold"
                          : isLiburSem
                          ? "text-black font-bold"
                          : isTglMerah
                          ? "text-black font-bold"
                          : isBukanJadwal
                          ? "text-black font-bold"
                          : ""
                      }`}
                      style={{
                        backgroundColor: isSundayDay
                          ? customColors.hariMinggu
                          : isLiburSem
                          ? customColors.liburSemester
                          : isTglMerah
                          ? customColors.tanggalMerah
                          : isBukanJadwal
                          ? customColors.bukanJadwal
                          : customColors.jadwalMengajar,
                      }}
                    >
                      {stats.total > 0 ? stats.hadir : ""}
                    </td>
                  );
                })}
                <td
                  className="border px-2 py-1 text-xs text-center"
                  colSpan={4}
                >
                  -
                </td>
              </tr>

              {/* Baris Persen Hadir per Tanggal */}
              <tr className="bg-green-50 font-semibold">
                <td
                  className="freeze-no border px-2 py-1 text-xs text-center"
                  colSpan={3}
                >
                  % Hadir
                </td>
                {Array.from({ length: daysInMonth }, (_, day) => {
                  const dayNum = day + 1;
                  const stats = getAttendanceByDate()[dayNum] || {
                    hadir: 0,
                    total: 0,
                  };
                  const percentage =
                    stats.total > 0
                      ? ((stats.hadir / stats.total) * 100).toFixed(0) + "%"
                      : "";
                  const isTglMerah = isTanggalMerah(dayNum);
                  const isSundayDay = isSunday(dayNum);
                  const isLiburSem = isLiburSemester(dayNum); // TAMBAHAN BARU
                  const isBukanJadwal = isBukanJadwalMengajar(dayNum);

                  return (
                    <td
                      key={day}
                      className={`border px-1 py-1 text-sm ${
                        isSundayDay
                          ? "text-black font-bold"
                          : isLiburSem
                          ? "text-black font-bold"
                          : isTglMerah
                          ? "text-black font-bold"
                          : isBukanJadwal
                          ? "text-black font-bold"
                          : ""
                      }`}
                      style={{
                        backgroundColor: isSundayDay
                          ? customColors.hariMinggu
                          : isLiburSem
                          ? customColors.liburSemester
                          : isTglMerah
                          ? customColors.tanggalMerah
                          : isBukanJadwal
                          ? customColors.bukanJadwal
                          : customColors.jadwalMengajar,
                      }}
                    >
                      {percentage}
                    </td>
                  );
                })}
                <td
                  className="border px-2 py-1 text-xs text-center"
                  colSpan={4}
                >
                  -
                </td>
              </tr>

              {/* BARIS BARU: Total Keseluruhan - UBAH BAGIAN INI */}
              <tr className="bg-yellow-50 font-bold">
                <td
                  className="freeze-no border px-2 py-1 text-xs text-center"
                  colSpan={3 + daysInMonth} // ← UBAH: Gabungkan kolom dari No sampai tanggal terakhir
                >
                  TOTAL
                </td>
                <td className="border px-2 py-1 text-xs text-center text-green-700">
                  {getTotalSummary().totalH}
                </td>
                <td className="border px-2 py-1 text-xs text-center text-blue-700">
                  {getTotalSummary().totalS}
                </td>
                <td className="border px-2 py-1 text-xs text-center text-yellow-700">
                  {getTotalSummary().totalI}
                </td>
                <td className="border px-2 py-1 text-xs text-center text-red-700">
                  {getTotalSummary().totalA}
                </td>
              </tr>

              {/* BARIS BARU: Persentase Keseluruhan - UBAH BAGIAN INI */}
              <tr className="bg-orange-50 font-bold">
                <td
                  className="freeze-no border px-2 py-1 text-xs text-center"
                  colSpan={3 + daysInMonth} // ← UBAH: Gabungkan kolom dari No sampai tanggal terakhir
                >
                  PERSENTASE BULANAN
                </td>
                <td className="border px-2 py-1 text-xs text-center text-green-700">
                  {getTotalSummary().percentH}%
                </td>
                <td className="border px-2 py-1 text-xs text-center text-blue-700">
                  {getTotalSummary().percentS}%
                </td>
                <td className="border px-2 py-1 text-xs text-center text-yellow-700">
                  {getTotalSummary().percentI}%
                </td>
                <td className="border px-2 py-1 text-xs text-center text-red-700">
                  {getTotalSummary().percentA}%
                </td>
              </tr>

              {/* BARIS BARU: Hari Efektif */}
              <tr className="bg-purple-50 font-bold">
                <td
                  className="freeze-no border px-2 py-1 text-xs text-center"
                  colSpan={3 + daysInMonth}
                >
                  HARI EFEKTIF
                </td>
                <td
                  className="border px-2 py-1 text-xs text-center text-purple-700"
                  colSpan={4}
                >
                  {getHariEfektif()} Hari
                </td>
              </tr>
            </tfoot>
          </table>

          {/* TABEL BARU: Informasi Siswa */}
          <div className="mt-6">
            <table className="min-w-full border-collapse border border-gray-200">
              <thead>
                <tr className="bg-purple-100">
                  <th
                    className="border px-4 py-2 text-sm font-bold"
                    colSpan={3}
                  >
                    📊 Informasi Jumlah Siswa
                  </th>
                </tr>
              </thead>
              <tbody>
                <tr className="bg-purple-50">
                  <td className="border px-4 py-2 text-sm font-semibold text-center">
                    Laki-laki
                  </td>
                  <td className="border px-4 py-2 text-sm font-semibold text-center">
                    Perempuan
                  </td>
                  <td className="border px-4 py-2 text-sm font-semibold text-center">
                    Total Siswa
                  </td>
                </tr>
                <tr className="bg-white">
                  <td className="border px-4 py-2 text-center text-lg font-bold text-blue-700">
                    {getGenderSummary().lakiLaki}
                  </td>
                  <td className="border px-4 py-2 text-center text-lg font-bold text-pink-700">
                    {getGenderSummary().perempuan}
                  </td>
                  <td className="border px-4 py-2 text-center text-lg font-bold text-purple-700">
                    {getGenderSummary().total}
                  </td>
                </tr>
              </tbody>
            </table>
          </div>

          {/* TABEL KETERANGAN TANGGAL MERAH */}
          {filteredTanggalMerah.length > 0 && (
            <div className="mt-6">
              <table className="min-w-full border-collapse border border-gray-200">
                <thead>
                  <tr className="bg-red-100">
                    <th
                      className="border px-4 py-2 text-sm font-bold"
                      colSpan={2}
                    >
                      📅 Keterangan Tanggal Merah / Libur
                    </th>
                  </tr>
                  <tr className="bg-red-50">
                    <td className="border px-4 py-2 text-sm font-semibold text-center">
                      Tanggal
                    </td>
                    <td className="border px-4 py-2 text-sm font-semibold text-center">
                      Keterangan
                    </td>
                  </tr>
                </thead>
                <tbody>
                  {filteredTanggalMerah
                    .sort((a, b) => {
                      // Sort by date
                      const [dayA, monthA, yearA] = a.tanggal
                        .split("/")
                        .map(Number);
                      const [dayB, monthB, yearB] = b.tanggal
                        .split("/")
                        .map(Number);
                      return (
                        new Date(yearA, monthA - 1, dayA).getTime() -
                        new Date(yearB, monthB - 1, dayB).getTime()
                      );
                    })
                    .map((tm, index) => {
                      // TAMBAHKAN: Helper function untuk format rentang tanggal
                      const formatDateRange = (
                        startDate: string,
                        endDate?: string
                      ) => {
                        if (!endDate) return startDate;

                        const [dayStart, monthStart, yearStart] =
                          startDate.split("/");
                        const [dayEnd, monthEnd, yearEnd] = endDate.split("/");

                        // Jika bulan dan tahun sama
                        if (monthStart === monthEnd && yearStart === yearEnd) {
                          return `${dayStart} - ${dayEnd}/${monthStart}/${yearStart}`;
                        }
                        // Jika tahun sama tapi bulan beda
                        else if (yearStart === yearEnd) {
                          return `${dayStart}/${monthStart} - ${dayEnd}/${monthEnd}/${yearEnd}`;
                        }
                        // Jika tahun berbeda
                        else {
                          return `${startDate} - ${endDate}`;
                        }
                      };

                      return (
                        <tr
                          key={index}
                          className={index % 2 === 0 ? "bg-white" : "bg-red-50"}
                        >
                          <td className="border px-4 py-2 text-center text-sm">
                            {(() => {
                              // Helper untuk menampilkan tanggal yang relevan dengan bulan dipilih
                              if (!tm.tanggalAkhir) return tm.tanggal;

                              const [dayStart, monthStart, yearStart] =
                                tm.tanggal.split("/").map(Number);
                              const [dayEnd, monthEnd, yearEnd] =
                                tm.tanggalAkhir.split("/").map(Number);

                              // Jika rentang dalam bulan yang sama
                              if (
                                monthStart === selectedMonth &&
                                monthEnd === selectedMonth &&
                                yearStart === selectedYear &&
                                yearEnd === selectedYear
                              ) {
                                return formatDateRange(
                                  tm.tanggal,
                                  tm.tanggalAkhir
                                );
                              }

                              // Jika rentang melewati beberapa bulan
                              // Di bulan pertama: tampilkan "DD/MM/YYYY - Akhir Bulan"
                              if (
                                monthStart === selectedMonth &&
                                yearStart === selectedYear
                              ) {
                                const lastDay = new Date(
                                  selectedYear,
                                  selectedMonth,
                                  0
                                ).getDate();
                                return `${tm.tanggal} - ${String(
                                  lastDay
                                ).padStart(2, "0")}/${String(
                                  selectedMonth
                                ).padStart(2, "0")}/${selectedYear}`;
                              }

                              // Di bulan terakhir: tampilkan "Awal Bulan - DD/MM/YYYY"
                              if (
                                monthEnd === selectedMonth &&
                                yearEnd === selectedYear
                              ) {
                                return `01/${String(selectedMonth).padStart(
                                  2,
                                  "0"
                                )}/${selectedYear} - ${tm.tanggalAkhir}`;
                              }

                              // Di bulan tengah: tampilkan "01 - 31 (atau akhir bulan)"
                              const lastDay = new Date(
                                selectedYear,
                                selectedMonth,
                                0
                              ).getDate();
                              return `01/${String(selectedMonth).padStart(
                                2,
                                "0"
                              )}/${selectedYear} - ${String(lastDay).padStart(
                                2,
                                "0"
                              )}/${String(selectedMonth).padStart(
                                2,
                                "0"
                              )}/${selectedYear}`;
                            })()}
                          </td>
                          <td className="border px-4 py-2 text-sm">
                            {tm.deskripsi}
                          </td>
                        </tr>
                      );
                    })}
                </tbody>
              </table>
            </div>
          )}

          {Object.keys(editedRecords).length > 0 && (
            <div className="mt-6 flex justify-center gap-4">
              <button
                onClick={handleSaveChanges}
                disabled={isSaving}
                className={`px-6 py-2 rounded-lg font-medium text-white ${
                  isSaving
                    ? "bg-blue-400 cursor-not-allowed"
                    : "bg-blue-600 hover:bg-blue-700"
                }`}
              >
                {isSaving
                  ? "⏳ Menyimpan..."
                  : `💾 Simpan Perubahan (${
                      Object.keys(editedRecords).length
                    })`}
              </button>
              <button
                onClick={() => setEditedRecords({})}
                disabled={isSaving}
                className="px-6 py-2 bg-gray-500 hover:bg-gray-600 text-white rounded-lg font-medium"
              >
                ❌ Batal
              </button>
            </div>
          )}
          <div className="mt-6 flex justify-center gap-4">
            <button
              onClick={downloadPDF}
              className="px-1 py-0.5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg font-medium transition-colors"
            >
              📄 Download PDF
            </button>
            <button
              onClick={handleDeleteAllAttendance}
              disabled={isDeleting}
              className={`px-6 py-2 text-white rounded-lg font-medium ${
                isDeleting
                  ? "bg-red-400 cursor-not-allowed"
                  : "bg-red-600 hover:bg-red-700"
              }`}
            >
              {isDeleting ? "Memproses..." : "🗑️ Hapus Semua Data Absensi"}
            </button>
          </div>
        </div>
      </div>
      {/* Modal Detail Siswa */}
      {showModal && selectedStudent && (
        <div
          className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50"
          onClick={closeModal}
        >
          <div
            className="bg-white rounded-lg shadow-xl p-6 max-w-md w-full mx-4"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex justify-between items-center mb-4">
              <h3 className="text-xl font-bold text-gray-800">Detail Siswa</h3>
              <button
                onClick={closeModal}
                className="text-gray-400 hover:text-gray-600 text-2xl font-bold"
              >
                ×
              </button>
            </div>

            <div className="space-y-3 mb-6">
              <div className="flex border-b pb-2">
                <span className="font-semibold text-gray-600 w-24">NISN:</span>
                <span className="text-gray-800">
                  {selectedStudent.nisn || "N/A"}
                </span>
              </div>
              <div className="flex border-b pb-2">
                <span className="font-semibold text-gray-600 w-24">Nama:</span>
                <span className="text-gray-800">
                  {selectedStudent.name || "N/A"}
                </span>
              </div>
              <div className="flex border-b pb-2">
                <span className="font-semibold text-gray-600 w-24">Kelas:</span>
                <span className="text-gray-800">
                  {selectedStudent.kelas || "N/A"}
                </span>
              </div>
            </div>

            <div className="flex gap-3">
              <button
                onClick={() => {
                  handleDeleteStudentAttendance(selectedStudent);
                  closeModal();
                }}
                disabled={deletingStudentId === selectedStudent.id || isSaving}
                className={`flex-1 px-4 py-2 rounded-lg font-medium transition-colors ${
                  deletingStudentId === selectedStudent.id || isSaving
                    ? "bg-gray-300 text-gray-500 cursor-not-allowed"
                    : "bg-red-500 hover:bg-red-600 text-white"
                }`}
              >
                {deletingStudentId === selectedStudent.id
                  ? "⏳ Menghapus..."
                  : "🗑️ Hapus Riwayat Absensi"}
              </button>
              <button
                onClick={closeModal}
                className="px-4 py-2 bg-gray-500 hover:bg-gray-600 text-white rounded-lg font-medium transition-colors"
              >
                Tutup
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

const TanggalMerahTab: React.FC<{
  onRefresh: () => void;
}> = ({ onRefresh }) => {
  // TAMBAHKAN DATA LIBUR NASIONAL DI SINI
  const liburNasionalOptions = [
    "Tahun Baru Masehi",
    "Tahun Baru Imlek",
    "Hari Raya Nyepi",
    "Wafat Isa Al-Masih",
    "Hari Buruh Internasional",
    "Kenaikan Isa Al-Masih",
    "Hari Raya Waisak",
    "Hari Lahir Pancasila",
    "Libur Awal Puasa Ramadhan",
    "Libur Hari Raya Idul Fitri",
    "Hari Raya Idul Adha",
    "Tahun Baru Islam",
    "Maulid Nabi Muhammad SAW",
    "Hari Kemerdekaan RI",
    "Isra Mikraj Nabi Muhammad SAW",
    "Hari Raya Natal",
    "Cuti Bersama Lebaran",
    "Cuti Bersama Tahun Baru",
    "Cuti Bersama Natal",
    "Libur Akhir Semester Ganjil",
    "Libur Akhir Semester Genap",
  ];
  const [tanggalMerahList, setTanggalMerahList] = useState<TanggalMerah[]>([]);
  const [tanggal, setTanggal] = useState("");
  const [tanggalAkhir, setTanggalAkhir] = useState("");
  const [deskripsi, setDeskripsi] = useState("");
  const [loading, setLoading] = useState<boolean>(true);
  const [isSaving, setIsSaving] = useState<boolean>(false);
  const [isDeleting, setIsDeleting] = useState<boolean>(false);
  const [editingIndex, setEditingIndex] = useState<number | null>(null);
  const [showDropdown, setShowDropdown] = useState<boolean>(false);
  const [searchQuery, setSearchQuery] = useState<string>("");
  const dropdownRef = useRef<HTMLDivElement>(null);

  // TAMBAHKAN: Helper untuk cek apakah perlu rentang tanggal
  const needsDateRange = React.useMemo(() => {
    const lowerDesc = deskripsi.toLowerCase();
    return (
      lowerDesc.includes("libur akhir semester") ||
      lowerDesc.includes("libur awal puasa ramadhan") ||
      lowerDesc.includes("libur hari raya idul fitri")
    );
  }, [deskripsi]);

  // Fetch data tanggal merah
  useEffect(() => {
    fetchTanggalMerah();
  }, []);

  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (
        dropdownRef.current &&
        !dropdownRef.current.contains(event.target as Node)
      ) {
        setShowDropdown(false);
      }
    };

    document.addEventListener("mousedown", handleClickOutside);
    return () => {
      document.removeEventListener("mousedown", handleClickOutside);
    };
  }, []);

  const fetchTanggalMerah = () => {
    setLoading(true);
    fetch(`${endpoint}?action=tanggalMerah`)
      .then((res) => {
        if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
        return res.json();
      })
      .then((data) => {
        if (data.success) {
          setTanggalMerahList(data.data || []);
        } else {
          alert("❌ Gagal memuat data tanggal merah: " + data.message);
          setTanggalMerahList([]);
        }
        setLoading(false);
      })
      .catch((error) => {
        console.error("Error fetch:", error);
        alert("❌ Gagal memuat data tanggal merah. Cek console untuk detail.");
        setLoading(false);
      });
  };

  const filteredOptions = React.useMemo(() => {
    if (!searchQuery.trim()) return liburNasionalOptions;
    return liburNasionalOptions.filter((option) =>
      option.toLowerCase().includes(searchQuery.toLowerCase())
    );
  }, [searchQuery, liburNasionalOptions]);

  const handleSelectOption = (option: string) => {
    setDeskripsi(option);
    setSearchQuery("");
    setShowDropdown(false);
  };

  const handleSubmit = () => {
    if (!tanggal || !deskripsi) {
      alert("⚠️ Tanggal dan Deskripsi wajib diisi!");
      return;
    }

    // TAMBAHKAN: Validasi untuk libur semester
    if (needsDateRange && !tanggalAkhir) {
      alert("⚠️ Tanggal akhir wajib diisi untuk Libur Semester!");
      return;
    }

    setIsSaving(true);

    fetch(endpoint, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        type: "tanggalMerah",
        tanggal: formatDateDDMMYYYY(tanggal),
        tanggalAkhir: tanggalAkhir ? formatDateDDMMYYYY(tanggalAkhir) : "", // TAMBAHKAN
        deskripsi,
      }),
    })
      .then(() => {
        alert("✅ Tanggal merah berhasil ditambahkan!");
        setTanggal("");
        setTanggalAkhir(""); // TAMBAHKAN
        setDeskripsi("");
        fetchTanggalMerah();
        onRefresh();
        setIsSaving(false);
      })
      .catch(() => {
        alert("❌ Gagal menambahkan tanggal merah.");
        setIsSaving(false);
      });
  };

  const handleEdit = (index: number) => {
    const item = tanggalMerahList[index];
    const [day, month, year] = item.tanggal.split("/");
    setTanggal(`${year}-${month}-${day}`);

    // TAMBAHKAN: Set tanggal akhir jika ada
    if (item.tanggalAkhir) {
      const [dayEnd, monthEnd, yearEnd] = item.tanggalAkhir.split("/");
      setTanggalAkhir(`${yearEnd}-${monthEnd}-${dayEnd}`);
    } else {
      setTanggalAkhir("");
    }

    setDeskripsi(item.deskripsi);
    setEditingIndex(index);
  };

  const handleUpdate = () => {
    if (!tanggal || !deskripsi || editingIndex === null) {
      alert("⚠️ Tanggal dan Deskripsi wajib diisi!");
      return;
    }

    // TAMBAHKAN: Validasi untuk libur semester
    if (needsDateRange && !tanggalAkhir) {
      alert("⚠️ Tanggal akhir wajib diisi untuk Libur Semester!");
      return;
    }

    const oldItem = tanggalMerahList[editingIndex];
    setIsSaving(true);

    fetch(endpoint, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        type: "editTanggalMerah",
        tanggalLama: oldItem.tanggal,
        tanggalBaru: formatDateDDMMYYYY(tanggal),
        tanggalAkhir: tanggalAkhir ? formatDateDDMMYYYY(tanggalAkhir) : "", // TAMBAHKAN
        deskripsi,
      }),
    })
      .then(() => {
        alert("✅ Tanggal merah berhasil diperbarui!");
        setTanggal("");
        setTanggalAkhir(""); // TAMBAHKAN
        setDeskripsi("");
        setEditingIndex(null);
        fetchTanggalMerah();
        onRefresh();
        setIsSaving(false);
      })
      .catch(() => {
        alert("❌ Gagal memperbarui tanggal merah.");
        setIsSaving(false);
      });
  };

  const handleDelete = (tanggalToDelete: string) => {
    if (!confirm(`Yakin ingin menghapus tanggal merah: ${tanggalToDelete}?`)) {
      return;
    }

    setIsDeleting(true);

    fetch(endpoint, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        type: "deleteTanggalMerah",
        tanggal: tanggalToDelete,
      }),
    })
      .then(() => {
        alert("✅ Tanggal merah berhasil dihapus!");
        fetchTanggalMerah();
        onRefresh();
        setIsDeleting(false);
      })
      .catch(() => {
        alert("❌ Gagal menghapus tanggal merah.");
        setIsDeleting(false);
      });
  };

  const handleCancel = () => {
    setTanggal("");
    setTanggalAkhir(""); // TAMBAHKAN
    setDeskripsi("");
    setEditingIndex(null);
  };

  if (loading) {
    return (
      <div className="text-center py-8">
        <p className="text-gray-500">Memuat data tanggal merah...</p>
      </div>
    );
  }

  return (
    <div className="max-w-4xl mx-auto" style={{ paddingBottom: "70px" }}>
      <div className="bg-white p-6 rounded-lg shadow-md mb-6">
        <h2 className="text-xl font-bold mb-4 text-center text-blue-600">
          {editingIndex !== null
            ? "Edit Tanggal Merah"
            : "Tambah Tanggal Merah"}
        </h2>
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
          <input
            type="date"
            value={tanggal}
            onChange={(e) => setTanggal(e.target.value)}
            className="w-full border border-gray-300 px-4 py-2 rounded-lg"
            disabled={isSaving}
            placeholder="Tanggal Mulai"
          />
          {/* TAMBAHKAN: Input tanggal akhir yang muncul conditional */}
          {needsDateRange && (
            <input
              type="date"
              value={tanggalAkhir}
              onChange={(e) => setTanggalAkhir(e.target.value)}
              className="w-full border border-gray-300 px-4 py-2 rounded-lg"
              disabled={isSaving}
              placeholder="Tanggal Akhir (untuk Libur Semester)"
            />
          )}
          {/* ⬇️ GANTI DENGAN KODE INI ⬇️ */}
          <div className="relative" ref={dropdownRef}>
            <input
              type="text"
              placeholder="Deskripsi (cth: Hari Raya Nyepi)"
              value={deskripsi}
              onChange={(e) => {
                setDeskripsi(e.target.value);
                setSearchQuery(e.target.value);
                setShowDropdown(true);
              }}
              onFocus={() => setShowDropdown(true)}
              className="w-full border border-gray-300 px-4 py-2 rounded-lg"
              disabled={isSaving}
            />

            {showDropdown && filteredOptions.length > 0 && (
              <div className="absolute z-50 w-full mt-1 bg-white border border-gray-300 rounded-lg shadow-lg max-h-60 overflow-y-auto">
                {filteredOptions.map((option, index) => (
                  <div
                    key={index}
                    onClick={() => handleSelectOption(option)}
                    className="px-4 py-2 hover:bg-blue-50 cursor-pointer text-sm transition-colors"
                  >
                    {option}
                  </div>
                ))}
              </div>
            )}

            {showDropdown && searchQuery && filteredOptions.length === 0 && (
              <div className="absolute z-50 w-full mt-1 bg-white border border-gray-300 rounded-lg shadow-lg p-4 text-center text-gray-500 text-sm">
                Tidak ada hasil untuk "{searchQuery}"
              </div>
            )}
          </div>
        </div>

        {/* TAMBAHKAN: Info helper untuk libur semester */}
        {needsDateRange && (
          <div className="mb-4 p-3 bg-blue-50 border border-blue-200 rounded-lg">
            <p className="text-sm text-blue-700">
              💡{" "}
              <strong>
                {deskripsi.toLowerCase().includes("libur akhir semester")
                  ? "Libur Semester"
                  : deskripsi
                      .toLowerCase()
                      .includes("libur awal puasa ramadhan")
                  ? "Libur Awal Puasa Ramadhan"
                  : "Libur Hari Raya Idul Fitri"}
                :
              </strong>{" "}
              Silakan pilih tanggal mulai dan tanggal akhir libur.
            </p>
          </div>
        )}
        <div className="flex justify-center gap-4">
          {editingIndex !== null ? (
            <>
              <button
                onClick={handleUpdate}
                disabled={isSaving}
                className={`px-6 py-2 rounded-lg font-medium transition-colors ${
                  isSaving
                    ? "bg-blue-400 cursor-not-allowed"
                    : "bg-blue-600 hover:bg-blue-700"
                } text-white`}
              >
                {isSaving ? "⏳ Menyimpan..." : "💾 Update"}
              </button>
              <button
                onClick={handleCancel}
                disabled={isSaving}
                className="px-6 py-2 bg-gray-500 hover:bg-gray-600 text-white rounded-lg font-medium"
              >
                ❌ Batal
              </button>
            </>
          ) : (
            <button
              onClick={handleSubmit}
              disabled={isSaving}
              className={`px-6 py-2 rounded-lg font-medium transition-colors ${
                isSaving
                  ? "bg-blue-400 cursor-not-allowed"
                  : "bg-blue-600 hover:bg-blue-700"
              } text-white`}
            >
              {isSaving ? "⏳ Menyimpan..." : "➕ Tambah Tanggal Merah"}
            </button>
          )}
        </div>
      </div>

      <div className="bg-white p-6 rounded-lg shadow-md">
        <h3 className="text-lg font-semibold text-gray-700 mb-4">
          Daftar Tanggal Merah ({tanggalMerahList.length})
        </h3>
        {tanggalMerahList.length === 0 ? (
          <p className="text-center text-gray-500 py-8">
            Belum ada data tanggal merah.
          </p>
        ) : (
          <div className="space-y-3">
            {tanggalMerahList.map((item, index) => (
              <div
                key={index}
                className="flex justify-between items-center bg-gray-50 border border-gray-200 px-4 py-3 rounded-lg"
              >
                <div>
                  <p className="font-medium text-gray-800">
                    {item.tanggalAkhir
                      ? `${item.tanggal} - ${item.tanggalAkhir}`
                      : item.tanggal}
                  </p>
                  <p className="text-sm text-gray-600">{item.deskripsi}</p>
                </div>
                <div className="flex gap-2">
                  <button
                    onClick={() => handleEdit(index)}
                    disabled={isSaving || isDeleting}
                    className={`text-xs px-3 py-1 rounded transition-colors ${
                      isSaving || isDeleting
                        ? "bg-yellow-400 cursor-not-allowed"
                        : "bg-yellow-500 hover:bg-yellow-600"
                    } text-white`}
                  >
                    ✏️ Edit
                  </button>
                  <button
                    onClick={() => handleDelete(item.tanggal)}
                    disabled={isSaving || isDeleting}
                    className={`text-xs px-3 py-1 rounded transition-colors ${
                      isSaving || isDeleting
                        ? "bg-red-400 cursor-not-allowed"
                        : "bg-red-500 hover:bg-red-600"
                    } text-white`}
                  >
                    🗑️ Hapus
                  </button>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
};

export default StudentAttendanceApp;
