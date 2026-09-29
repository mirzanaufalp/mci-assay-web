import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { ScatterChart, Scatter, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, ReferenceLine, Cell } from 'recharts';

const ZONE_MCI = { "tp": 10.0, "soft": 30.0, "hard cluster": 65.0, "hard": 90.0 };
const H_SCALE_M = 15.0;
const V_SCALE_M = 0.40;
const MAX_ASSAY_WEIGHT = 0.95;
const MIN_EFFECTIVE_WEIGHT = 0.01;

// Fungsi gradasi warna mirip colormap matplotlib
const getContourColor = (mci) => {
  if (mci <= 15) return "#3b1443";       // Ungu sangat tua
  if (mci <= 30) return "#453781";       // Ungu gelap
  if (mci <= 45) return "#33638d";       // Biru gelap
  if (mci <= 60) return "#20a387";       // Toska / Hijau kebiruan
  if (mci <= 75) return "#35b779";       // Hijau cerah
  if (mci <= 90) return "#90d743";       // Kuning kehijauan
  return "#fde725";                      // Kuning terang
};

export default function App() {
  const [logs, setLogs] = useState([]);
  const [chartData, setChartData] = useState([]);
  const [fullExportData, setFullExportData] = useState([]);
  const [assayLines, setAssayLines] = useState([]);
  const [isProcessing, setIsProcessing] = useState(false);

  const addLog = (msg) => setLogs(prev => [...prev, msg]);

  const handleFileUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    setIsProcessing(true);
    setLogs(["Membaca file Excel..."]);
    setChartData([]);
    setFullExportData([]);
    setAssayLines([]);

    try {
      const data = await file.arrayBuffer();
      const wb = XLSX.read(data);

      if (!wb.Sheets["res_L3"] || !wb.Sheets["Assay L3"]) {
         addLog("Error: Sheet 'res_L3' atau 'Assay L3' tidak ditemukan!");
         setIsProcessing(false);
         return;
      }

      const resRaw = XLSX.utils.sheet_to_json(wb.Sheets["res_L3"]);
      const assayRaw = XLSX.utils.sheet_to_json(wb.Sheets["Assay L3"]);

      let sumX = 0, sumY = 0;
      resRaw.forEach(r => { sumX += r.X; sumY += r.Y; });
      const cx = sumX / resRaw.length;
      const cy = sumY / resRaw.length;

      let covXX = 0, covXY = 0, covYY = 0;
      resRaw.forEach(r => {
         covXX += (r.X - cx) * (r.X - cx);
         covXY += (r.X - cx) * (r.Y - cy);
         covYY += (r.Y - cy) * (r.Y - cy);
      });
      
      const theta = 0.5 * Math.atan2(2 * covXY, covXX - covYY);
      const axisX = Math.cos(theta);
      const axisY = Math.sin(theta);

      let minS = Infinity;
      const resWithS = resRaw.map(r => {
         const s = (r.X - cx) * axisX + (r.Y - cy) * axisY;
         if (s < minS) minS = s;
         return { ...r, sRaw: s };
      }).map(r => ({ ...r, s: r.sRaw - minS }));

      const resValues = resWithS.map(r => r.RES).filter(v => v > 0).sort((a,b) => a-b);
      const p5 = resValues[Math.floor(resValues.length * 0.05)];
      const p95 = resValues[Math.floor(resValues.length * 0.95)];
      const logP5 = Math.log10(p5);
      const logP95 = Math.log10(p95);

      const resProcessed = resWithS.map(r => {
         if (r.RES <= 0) return { ...r, mciBase: null };
         let mci = 100 * ((Math.log10(r.RES) - logP5) / (logP95 - logP5));
         mci = Math.max(0, Math.min(100, mci));
         return { ...r, mciBase: mci };
      });

      const validAssays = [];
      assayRaw.forEach(a => {
         const zone = (a["Zone Modifikasofti"] || "").toString().trim().toLowerCase();
         const target = ZONE_MCI[zone];
         if (target !== undefined && a.XCollar && a.YCollar && a.ZCollar && a.From !== undefined && a.To !== undefined) {
             const sRaw = (a.XCollar - cx) * axisX + (a.YCollar - cy) * axisY;
             const s = sRaw - minS;
             const zMid = a.ZCollar - ((a.From + a.To) / 2.0);
             validAssays.push({
                 bhid: a.BHID, s, zMid, target,
                 zFrom: a.ZCollar - a.From,
                 zTo: a.ZCollar - a.To
             });
         }
      });
      setAssayLines(validAssays);

      addLog("Menerapkan constraint assay...");
      const rawExportList = [];
      const finalData = resProcessed.map(r => {
         if (r.mciBase === null) return null;
         let wSum = 0, wMax = 0, assayTargetSum = 0;

         validAssays.forEach(a => {
             const dh = r.s - a.s;
             const dv = r.Z - a.zMid;
             const w = Math.exp(-0.5 * (Math.pow(dh / H_SCALE_M, 2) + Math.pow(dv / V_SCALE_M, 2)));
             if (w >= MIN_EFFECTIVE_WEIGHT) {
                 wSum += w;
                 assayTargetSum += w * a.target;
                 if (w > wMax) wMax = w;
             }
         });

         let finalMCI = r.mciBase;
         let constraintWeight = 0;
         let localAssayTarget = "";
         if (wSum > 0) {
             localAssayTarget = assayTargetSum / wSum;
             constraintWeight = Math.min(MAX_ASSAY_WEIGHT, MAX_ASSAY_WEIGHT * wMax);
             finalMCI = (1.0 - constraintWeight) * r.mciBase + constraintWeight * localAssayTarget;
         }

         let mciClass = "Sedang";
         if (finalMCI <= 20) mciClass = "Sangat rendah";
         else if (finalMCI <= 40) mciClass = "Rendah";
         else if (finalMCI <= 60) mciClass = "Sedang";
         else if (finalMCI <= 80) mciClass = "Tinggi";
         else mciClass = "Sangat tinggi";

         rawExportList.push({
             X: r.X,
             Y: r.Y,
             Z: r.Z,
             RES: r.RES,
             LINE: r.LINE || "L3",
             Profile_m: Number(r.s.toFixed(4)),
             MCI_RES_Base: Number(r.mciBase.toFixed(4)),
             MCI_Assay_Target_Local: localAssayTarget !== "" ? Number(localAssayTarget.toFixed(4)) : "",
             Constraint_Weight: Number(constraintWeight.toFixed(4)),
             MCI_AssayConstrained: Number(finalMCI.toFixed(4)),
             MCI_Class: mciClass
         });

         return {
             s: Number(r.s.toFixed(2)),
             z: Number(r.Z.toFixed(2)),
             mci: Number(finalMCI.toFixed(2)),
             color: getContourColor(finalMCI)
         };
      }).filter(r => r !== null);

      setChartData(finalData);
      setFullExportData(rawExportList);
      addLog(`Selesai! Berhasil memproses ${finalData.length} titik koordinat.`);

    } catch (err) {
      addLog(`Error: ${err.message}`);
    }
    setIsProcessing(false);
  };

  const downloadFullCSV = () => {
    if (fullExportData.length === 0) return;
    const headers = "X,Y,Z,RES,LINE,Profile_m,MCI_RES_Base,MCI_Assay_Target_Local,Constraint_Weight,MCI_AssayConstrained,MCI_Class\n";
    const rows = fullExportData.map(d => 
      `${d.X},${d.Y},${d.Z},${d.RES},${d.LINE},${d.Profile_m},${d.MCI_RES_Base},${d.MCI_Assay_Target_Local},${d.Constraint_Weight},${d.MCI_AssayConstrained},${d.MCI_Class}`
    ).join("\n");
    
    const blob = new Blob([headers + rows], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", "L3_Hasil_MCI_Constrained_Full.csv");
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  return (
    <div className="min-h-screen bg-gray-50 p-8 text-gray-800 font-sans">
      <div className="max-w-6xl mx-auto bg-white rounded-xl shadow-lg p-6 border border-gray-100">
        <h1 className="text-3xl font-bold text-slate-800 mb-2 border-b pb-4">MCI Assay-Constrained Visualizer</h1>
        <p className="text-sm text-slate-500 mb-6">Aplikasi untuk menghitung profil MCI berdasarkan data resistivity dan constraint pemboran.</p>

        <div className="mb-6 p-5 bg-blue-50 border border-blue-100 rounded-lg shadow-inner flex flex-col md:flex-row justify-between items-center gap-4">
            <div className="w-full">
                <label className="block text-sm font-bold text-blue-900 mb-2">Upload File Excel (.xlsx)</label>
                <input
                  type="file"
                  accept=".xlsx"
                  onChange={handleFileUpload}
                  className="block w-full text-sm text-slate-500 file:mr-4 file:py-2.5 file:px-4 file:rounded-md file:border-0 file:text-sm file:font-semibold file:bg-blue-600 file:text-white hover:file:bg-blue-700 cursor-pointer transition-colors"
                  disabled={isProcessing}
                />
            </div>
            {fullExportData.length > 0 && (
                <button
                  onClick={downloadFullCSV}
                  className="w-full md:w-auto bg-emerald-600 hover:bg-emerald-700 text-white font-semibold px-5 py-2.5 rounded-md shadow transition-colors text-sm whitespace-nowrap flex items-center justify-center gap-2"
                >
                  <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"></path></svg>
                  Download Data Lengkap (CSV)
                </button>
            )}
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-4 gap-6">
          <div className="lg:col-span-1 bg-slate-50 p-4 rounded-lg border border-slate-200">
             <h3 className="font-bold text-slate-700 mb-3 border-b pb-2">Log Proses:</h3>
             <ul className="list-disc pl-4 text-xs text-slate-600 space-y-1.5 font-mono">
               {logs.length === 0 ? <li>Belum ada aktivitas.</li> : logs.map((l, i) => <li key={i}>{l}</li>)}
             </ul>
          </div>

          <div className="lg:col-span-3 h-[520px] w-full bg-white border border-slate-200 rounded-lg p-4 shadow-sm flex">
             {chartData.length > 0 ? (
                 <>
                   <div className="flex-grow h-full">
                     <ResponsiveContainer width="100%" height="100%">
                        <ScatterChart margin={{ top: 20, right: 10, bottom: 20, left: 10 }}>
                          <CartesianGrid strokeDasharray="3 3" opacity={0.3} />
                          <XAxis dataKey="s" type="number" name="Profile" unit="m" tick={{fontSize: 11}} />
                          <YAxis dataKey="z" type="number" name="Elevasi" unit="m" domain={['auto', 'auto']} tick={{fontSize: 11}} />
                          <Tooltip cursor={{ strokeDasharray: '3 3' }} />
                          <Scatter name="Kontur MCI" data={chartData} shape="square">
                            {chartData.map((entry, index) => (
                              <Cell key={`cell-${index}`} fill={entry.color} />
                            ))}
                          </Scatter>
                          {assayLines.map((line, i) => (
                              <ReferenceLine key={i} segment={[{ x: line.s, y: line.zFrom }, { x: line.s, y: line.zTo }]} stroke="#111" strokeWidth={3} opacity={0.8} />
                          ))}
                        </ScatterChart>
                     </ResponsiveContainer>
                   </div>
                   {/* Colorbar / Legenda Warna */}
                   <div className="w-14 flex flex-col items-center justify-center pl-2 ml-2 border-l border-slate-200 text-[10px] text-slate-600">
                      <span className="font-bold mb-1">90+</span>
                      <div className="w-4 h-full rounded shadow-inner" style={{ background: 'linear-gradient(to top, #3b1443, #453781, #33638d, #20a387, #35b779, #90d743, #fde725)' }}></div>
                      <span className="font-bold mt-1">0</span>
                   </div>
                 </>
             ) : (
                 <div className="flex flex-col items-center justify-center h-full w-full text-slate-400">
                    <svg className="w-12 h-12 mb-3 text-slate-300" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z"></path></svg>
                    <span>Penampang visual akan muncul di sini.</span>
                 </div>
             )}
          </div>
        </div>
      </div>
    </div>
  );
}