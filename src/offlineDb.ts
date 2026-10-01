import { openDB, IDBPDatabase } from "idb";

let ENDPOINT = "";
export const configureSync = (url: string) => {
  ENDPOINT = url;
};

const DB_NAME = "absensi-cache";
const DB_VERSION = 1;
const STORES = [
  "students",
  "jadwal",
  "tanggalMerah",
  "fotoAbsen",
  "attendance",
  "meta",
];

let dbPromise: Promise<IDBPDatabase> | null = null;
const getDb = () => {
  if (!dbPromise) {
    dbPromise = openDB(DB_NAME, DB_VERSION, {
      upgrade(db) {
        db.createObjectStore("students", { keyPath: "nisn" });
        db.createObjectStore("jadwal", { keyPath: "kelas" });
        db.createObjectStore("tanggalMerah", { keyPath: "tanggal" });
        db.createObjectStore("fotoAbsen", { keyPath: "key" });
        const att = db.createObjectStore("attendance", { keyPath: "key" });
        att.createIndex("tanggal", "tanggal");
        db.createObjectStore("meta");
      },
    });
  }
  return dbPromise;
};

let inflight: Promise<string[]> | null = null;

export function syncAll(force = false): Promise<string[]> {
  if (inflight && !force) return inflight;
  inflight = doSync(force).finally(() => {
    inflight = null;
  });
  return inflight;
}

async function doSync(force: boolean): Promise<string[]> {
  const db = await getDb();
  const vers = force ? {} : (await db.get("meta", "vers")) || {};

  const res = await fetch(
    `${ENDPOINT}?action=sync&v=${encodeURIComponent(JSON.stringify(vers))}`
  );
  if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
  const json = await res.json();
  if (!json.success) throw new Error(json.message || "Sync gagal");

  const ch = json.changes || {};
  const tx = db.transaction(STORES, "readwrite");

  const replaceAll = async (
    store: string,
    rows: any[],
    mapRow: (r: any, i: number) => any
  ) => {
    const s = tx.objectStore(store);
    await s.clear();
    for (let i = 0; i < rows.length; i++) await s.put(mapRow(rows[i], i));
  };

  if (ch.students)
    await replaceAll("students", ch.students, (r, i) => ({
      ...r,
      nisn: String(r.nisn).trim(),
      id: String(r.nisn).trim(),
      _i: i,
    }));

  if (ch.jadwal)
    await replaceAll("jadwal", ch.jadwal, (r, i) => ({
      ...r,
      kelas: String(r.kelas).trim(),
      _i: i,
    }));

  if (ch.tanggalMerah)
    await replaceAll("tanggalMerah", ch.tanggalMerah, (r, i) => ({
      ...r,
      _i: i,
    }));

  if (ch.fotoAbsen)
    await replaceAll("fotoAbsen", ch.fotoAbsen, (r) => ({
      ...r,
      key: `${r.tanggal}|${String(r.kelas).trim()}`,
    }));

  if (ch.absensi) {
    const s = tx.objectStore("attendance");
    if (ch.absensi.full) await s.clear();
    for (const [tgl, rows] of Object.entries<any[]>(ch.absensi.dates || {})) {
      if (!ch.absensi.full) {
        let c = await s.index("tanggal").openCursor(tgl);
        while (c) {
          await c.delete();
          c = await c.continue();
        }
      }
      for (const r of rows)
        await s.put({ ...r, key: `${tgl}|${String(r.nisn).trim()}` });
    }
  }

  await tx.objectStore("meta").put(json.vers, "vers");
  await tx.done;

  return Object.keys(ch);
}

const byOrder = (a: any, b: any) => (a._i ?? 0) - (b._i ?? 0);
const stripOrder = ({ _i, ...rest }: any) => rest;

export async function getStudentsLocal(): Promise<any[]> {
  const db = await getDb();
  return (await db.getAll("students")).sort(byOrder).map(stripOrder);
}

export async function getJadwalLocal(): Promise<any[]> {
  const db = await getDb();
  return (await db.getAll("jadwal")).sort(byOrder).map(stripOrder);
}

export async function getTanggalMerahLocal(): Promise<any[]> {
  const db = await getDb();
  return (await db.getAll("tanggalMerah")).sort(byOrder).map(stripOrder);
}

export async function getAttendanceByDate(tanggal: string): Promise<any[]> {
  const db = await getDb();
  return db.getAllFromIndex("attendance", "tanggal", tanggal);
}

export async function getAttendanceAllLocal(): Promise<any[]> {
  const db = await getDb();
  return db.getAll("attendance");
}

export async function getFotoAbsenLocal(
  tanggal: string,
  kelas: string
): Promise<string | null> {
  const db = await getDb();
  const row = await db.get("fotoAbsen", `${tanggal}|${kelas}`);
  return row?.fotoUrl || null;
}

export async function resetCache() {
  const db = await getDb();
  const tx = db.transaction(STORES, "readwrite");
  await Promise.all(STORES.map((s) => tx.objectStore(s).clear()));
  await tx.done;
}
