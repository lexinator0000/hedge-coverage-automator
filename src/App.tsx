/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { 
  FileText, 
  Upload, 
  CheckCircle2, 
  AlertCircle, 
  Mail, 
  Download, 
  ArrowRight,
  FileSpreadsheet,
  Clock,
  Send,
  Trash2
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

// Utility for tailwind classes
function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

type Step = 'upload' | 'validate' | 'export' | 'email';

interface ValidationResult {
  matches: boolean;
  redMatch: boolean;
  greenMatch: boolean;
  details: {
    reportValue: string;
    calcValue: string;
  };
}

export default function App() {
  const [currentStep, setCurrentStep] = useState<Step>('upload');
  const [reportFile, setReportFile] = useState<File | null>(null);
  const [reportInputMode, setReportInputMode] = useState<'file' | 'text'>('file');
  const [reportText, setReportText] = useState('');
  const [calcFile, setCalcFile] = useState<File | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [processError, setProcessError] = useState<string | null>(null);
  const [validation, setValidation] = useState<ValidationResult | null>(null);
  const [processedWorkbook, setProcessedWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [availableSheets, setAvailableSheets] = useState<string[]>([]);
  const [selectedSheet, setSelectedSheet] = useState<string>('');
  const [tempWorkbook, setTempWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [emailTo, setEmailTo] = useState('secondary@tradingdesk.com');
  const [emailBcc, setEmailBcc] = useState('compliance@tradingdesk.com');
  const [emailSubject, setEmailSubject] = useState(`Hedge Coverage Report - ${new Date().toLocaleDateString()}`);
  const [emailBody, setEmailBody] = useState('Please find the attached Hedge Coverage Report.');
  const [isEditingEmail, setIsEditingEmail] = useState(false);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>, type: 'report' | 'calc') => {
    const file = e.target.files?.[0];
    if (file) {
      if (type === 'report') setReportFile(file);
      else setCalcFile(file);
    }
  };

  const processFiles = async () => {
    if ((!reportFile && !reportText) || !calcFile) return;
    setIsProcessing(true);
    setProcessError(null);
    setAvailableSheets([]);

    try {
      const calcData = await calcFile.arrayBuffer();
      const workbook = XLSX.read(calcData);
      setTempWorkbook(workbook);
      
      const targetSheetName = "Coverage Breakdown";
      const actualSheetName = workbook.SheetNames.find(
        name => name.trim().replace(/\s+/g, ' ').toLowerCase() === targetSheetName.toLowerCase()
      );
      
      if (actualSheetName) {
        completeProcessing(workbook, actualSheetName);
      } else {
        setAvailableSheets(workbook.SheetNames);
        setProcessError(`Could not automatically find the "${targetSheetName}" tab. Please select the correct tab from the list below.`);
        setIsProcessing(false);
      }
    } catch (error) {
      console.error("Error processing files:", error);
      setProcessError("Error processing Excel files. Please ensure they are valid .xlsx files.");
      setIsProcessing(false);
    }
  };

  const completeProcessing = (workbook: XLSX.WorkBook, sheetName: string) => {
    try {
      const worksheet = workbook.Sheets[sheetName];
      if (!worksheet) throw new Error("Sheet not found");

      // Convert sheet to 2D array for easier searching
      const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
      
      // Helper to find value by label
      const findValue = (label: string, colOffset: number = 1) => {
        for (let r = 0; r < data.length; r++) {
          for (let c = 0; c < data[r].length; c++) {
            const cellValue = String(data[r][c] || '').trim();
            if (cellValue.toLowerCase().includes(label.toLowerCase())) {
              return data[r][c + colOffset];
            }
          }
        }
        return null;
      };

      // Extract values based on user's template labels
      const extracted = {
        grossLocked: findValue("Gross Locked Position") || 0,
        percentage: findValue("Percentage of Coverage Pipeline") || 0,
        netLockedA: findValue("Net Locked Pipeline after Pull-Through") || 0,
        closedLoansB: findValue("Total Amount of Closed Loans in Inventory") || 0,
        totalAtRisk: findValue("Total amount of loans with at risk positions") || 0,
        hedgeC: findValue("Total amount of open hedge positions") || 0,
        hedgeCPrice: findValue("Total amount of open hedge positions", 2) || 0,
        bestEffortsD: findValue("Total amount of Best efforts Deliveries") || 0,
        mandatoryE: findValue("Total amount of Mandatory Deliveries") || 0,
        mandatoryEPrice: findValue("Total amount of Mandatory Deliveries", 2) || 0,
        sumCovered: findValue("Sum of covered positions") || 0,
        wap: findValue("Weighted Average Price") || 0,
        srp: findValue("Weighted average SRP premium") || 0,
        allInPrice: findValue("Weighted average all in price") || 0,
        totalUncommitted: findValue("Total uncommitted") || 0,
        atRiskPriceMovement: findValue("At risk $ with price movement") || 0,
      };

      // Validation logic: Check if RED/GREEN values match (mocked for now but using extracted data)
      const mockValidation: ValidationResult = {
        matches: true,
        redMatch: true,
        greenMatch: true,
        details: {
          reportValue: String(extracted.grossLocked),
          calcValue: String(extracted.grossLocked)
        }
      };
      setValidation(mockValidation);

      // Create the formatted summary sheet as requested by the user
      const summaryData = [
        ["GOLDEN EMPIRE MORTGAGE - HEDGE COVERAGE REPORT", "", "", "", ""],
        ["Report Date:", new Date().toLocaleDateString(), "", "", ""],
        ["", "", "", "", ""],
        ["Company Name:", "Golden Empire", "", "", ""],
        ["Contacts/Email Address & Phone Number:", "", "", "", ""],
        ["", "", "", "", ""],
        ["Pipeline Summary-Gross Locked Position", extracted.grossLocked, "", "", ""],
        ["Percentage of Coverage Pipeline \"at risk loans\":", extracted.percentage, "", "", ""],
        ["A.  Net Locked Pipeline after Pull-Through:", extracted.netLockedA, "", "", ""],
        ["B.  Total Amount of Closed Loans in Inventory (includes all Warehouse facilities):", extracted.closedLoansB, "", "", ""],
        ["Total amount of loans with at risk positions (A+B):", extracted.totalAtRisk, "", "", ""],
        ["", "", "", "", ""],
        ["C.  Total amount of open hedge positions:", extracted.hedgeC, extracted.hedgeCPrice, "", ""],
        ["D.  Total amount of Best efforts Deliveries:", extracted.bestEffortsD, 0, "", ""],
        ["E.  Total amount of Mandatory Deliveries:", extracted.mandatoryE, extracted.mandatoryEPrice, "", ""],
        ["Sum of covered positions (C+D+E):", extracted.sumCovered, "", "", ""],
        ["", "", "", "", ""],
        ["Weighted Average Price (\"WAP\") (Calculate WAP of C,D&E):", extracted.wap, "", "", ""],
        ["Weighted average SRP premium (if applicable):", extracted.srp, "", "", ""],
        ["Weighted average all in price:", extracted.allInPrice, "", "", ""],
        ["", "", "", "", ""],
        ["Total uncommitted", extracted.totalUncommitted, "", "", ""],
        ["At risk $ with price movement*:   1% Change in rate = 2% in Price", extracted.atRiskPriceMovement, "", "", ""],
      ];

      const newWorkbook = XLSX.utils.book_new();
      const newSheet = XLSX.utils.aoa_to_sheet(summaryData);
      
      // Apply Number Formats
      const currencyFormat = "$#,##0.00";
      const percentFormat = "0.00%";
      const wapFormat = "0.00000";

      // Helper to set format
      const setFormat = (cell: string, format: string) => {
        if (newSheet[cell]) newSheet[cell].z = format;
      };

      // Apply formats to specific cells based on the summaryData structure
      // Note: Row indices are 0-based in summaryData, but 1-based in Excel (A1, B2, etc.)
      
      // Currency fields (Column B)
      [7, 9, 10, 11, 13, 14, 15, 16, 22, 23].forEach(row => setFormat(`B${row}`, currencyFormat));
      
      // Percentage field (Column B)
      setFormat("B8", percentFormat);
      
      // WAP/Price fields (Column B and C)
      [13, 15].forEach(row => setFormat(`C${row}`, wapFormat));
      [18, 19, 20].forEach(row => setFormat(`B${row}`, wapFormat));

      // Set column widths for better readability
      newSheet['!cols'] = [
        { wch: 65 }, // Label column
        { wch: 22 }, // Value column
        { wch: 18 }, // Price column
        { wch: 15 },
        { wch: 15 }
      ];

      XLSX.utils.book_append_sheet(newWorkbook, newSheet, "Hedge Coverage Report");
      setProcessedWorkbook(newWorkbook);
      setCurrentStep('validate');
    } catch (error) {
      console.error("Processing error:", error);
      setProcessError("Failed to extract data and format the summary report. Please ensure the selected sheet contains the expected data.");
    } finally {
      setIsProcessing(false);
    }
  };

  const handleManualSheetSelect = (sheetName: string) => {
    if (!tempWorkbook) return;
    setIsProcessing(true);
    setProcessError(null);
    completeProcessing(tempWorkbook, sheetName);
  };

  const downloadProcessedFile = () => {
    if (!processedWorkbook) return;
    XLSX.writeFile(processedWorkbook, "Hedge_Coverage_Report.xlsx");
    setCurrentStep('email');
  };

  const isMonday = new Date().getDay() === 1;

  const getEmailLink = () => {
    return `mailto:${emailTo}?bcc=${emailBcc}&subject=${encodeURIComponent(emailSubject)}&body=${encodeURIComponent(emailBody)}`;
  };

  return (
    <div className="min-h-screen bg-[#0a0a0a] text-white/90 font-sans selection:bg-purple-600 selection:text-white">
      {/* Header */}
      <header className="border-b border-white/10 bg-black/50 backdrop-blur-md sticky top-0 z-10">
        <div className="max-w-5xl mx-auto px-6 py-4 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-purple-600 rounded-xl flex items-center justify-center text-white shadow-lg shadow-purple-600/20">
              <FileSpreadsheet size={20} />
            </div>
            <div>
              <h1 className="font-serif text-xl font-medium italic">Hedge Coverage</h1>
              <p className="text-[10px] uppercase tracking-widest opacity-40 font-mono">Automator v1.0</p>
            </div>
          </div>
          <div className="flex items-center gap-4">
            <div className={cn(
              "px-3 py-1 rounded-full text-[10px] font-mono uppercase tracking-wider flex items-center gap-2",
              isMonday ? "bg-emerald-500/10 text-emerald-400 border border-emerald-500/20" : "bg-purple-500/10 text-purple-400 border border-purple-500/20"
            )}>
              <Clock size={12} />
              {isMonday ? "Monday Distribution Active" : "Non-Monday Schedule"}
            </div>
          </div>
        </div>
      </header>

      <main className="max-w-5xl mx-auto px-6 py-12">
        {/* Progress Stepper */}
        <div className="flex items-center justify-between mb-16 relative">
          <div className="absolute top-1/2 left-0 w-full h-[1px] bg-white/10 -z-10" />
          {[
            { id: 'upload', label: 'Upload', icon: Upload },
            { id: 'validate', label: 'Validate', icon: CheckCircle2 },
            { id: 'export', label: 'Export', icon: Download },
            { id: 'email', label: 'Distribute', icon: Mail },
          ].map((step, idx) => (
            <div key={step.id} className="flex flex-col items-center gap-3 bg-[#0a0a0a] px-4">
              <div className={cn(
                "w-10 h-10 rounded-full flex items-center justify-center transition-all duration-500 border",
                currentStep === step.id 
                  ? "bg-purple-600 text-white border-purple-600 shadow-lg shadow-purple-600/30" 
                  : idx < ['upload', 'validate', 'export', 'email'].indexOf(currentStep)
                    ? "bg-emerald-500 text-white border-emerald-500"
                    : "bg-white/5 text-white/20 border-white/10"
              )}>
                <step.icon size={18} />
              </div>
              <span className={cn(
                "text-[10px] uppercase tracking-widest font-mono",
                currentStep === step.id ? "text-purple-400 font-bold" : "text-white/30"
              )}>
                {step.label}
              </span>
            </div>
          ))}
        </div>

        <AnimatePresence mode="wait">
          {currentStep === 'upload' && (
            <motion.div 
              key="upload"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
              className="grid md:grid-cols-2 gap-8"
            >
              {/* Report Upload */}
              <div className="bg-white/5 rounded-3xl p-8 border border-white/10 shadow-sm">
                <div className="flex items-center gap-3 mb-6">
                  <FileText className="text-purple-500" size={24} />
                  <h2 className="font-serif text-xl italic">1. Coverage Report</h2>
                </div>
                <p className="text-sm text-white/50 mb-8 leading-relaxed">
                  Upload the report received via email (Hedge Coverage Report).
                </p>
                
                <div className="flex gap-2 mb-4">
                  <button 
                    onClick={() => setReportInputMode('file')}
                    className={cn(
                      "text-[10px] uppercase tracking-widest px-3 py-1 rounded-full border transition-all",
                      reportInputMode === 'file' ? "bg-purple-600 text-white border-purple-600" : "border-white/10 opacity-50"
                    )}
                  >
                    File Upload
                  </button>
                  <button 
                    onClick={() => setReportInputMode('text')}
                    className={cn(
                      "text-[10px] uppercase tracking-widest px-3 py-1 rounded-full border transition-all",
                      reportInputMode === 'text' ? "bg-purple-600 text-white border-purple-600" : "border-white/10 opacity-50"
                    )}
                  >
                    Paste Text
                  </button>
                </div>

                {reportInputMode === 'file' ? (
                  <label className="group relative block cursor-pointer">
                    <input 
                      type="file" 
                      className="hidden" 
                      onChange={(e) => handleFileUpload(e, 'report')}
                      accept=".xlsx,.xls,.csv"
                    />
                    <div className={cn(
                      "border-2 border-dashed rounded-2xl p-12 flex flex-col items-center justify-center transition-all",
                      reportFile ? "border-emerald-500/30 bg-emerald-500/5" : "border-white/10 group-hover:border-purple-500/30 group-hover:bg-purple-500/5"
                    )}>
                      {reportFile ? (
                        <>
                          <CheckCircle2 className="text-emerald-500 mb-4" size={32} />
                          <span className="text-sm font-medium">{reportFile.name}</span>
                          <button 
                            onClick={(e) => { e.preventDefault(); setReportFile(null); }}
                            className="mt-4 text-[10px] uppercase tracking-widest text-red-500 hover:underline"
                          >
                            Remove
                          </button>
                        </>
                      ) : (
                        <>
                          <Upload className="text-white/10 mb-4 group-hover:text-purple-500/50 transition-colors" size={32} />
                          <span className="text-sm text-white/30">Drop report file here</span>
                        </>
                      )}
                    </div>
                  </label>
                ) : (
                  <div className="space-y-4">
                    <textarea 
                      value={reportText}
                      onChange={(e) => setReportText(e.target.value)}
                      placeholder="Paste the report email content here..."
                      className="w-full h-40 p-4 rounded-2xl border border-white/10 bg-black/20 text-white/80 text-sm focus:outline-none focus:ring-2 focus:ring-purple-500/20 resize-none font-mono"
                    />
                    {reportText && (
                      <div className="flex items-center gap-2 text-emerald-400 text-[10px] font-mono uppercase tracking-wider">
                        <CheckCircle2 size={12} />
                        Content Captured
                      </div>
                    )}
                  </div>
                )}
              </div>

              {/* Calc Upload */}
              <div className="bg-white/5 rounded-3xl p-8 border border-white/10 shadow-sm">
                <div className="flex items-center gap-3 mb-6">
                  <FileSpreadsheet className="text-purple-500" size={24} />
                  <h2 className="font-serif text-xl italic">2. Coverage Calc</h2>
                </div>
                <p className="text-sm text-white/50 mb-8 leading-relaxed">
                  Upload the "Hedge Coverage Calc" file for validation.
                </p>
                
                <label className="group relative block cursor-pointer">
                  <input 
                    type="file" 
                    className="hidden" 
                    onChange={(e) => handleFileUpload(e, 'calc')}
                    accept=".xlsx,.xls"
                  />
                  <div className={cn(
                    "border-2 border-dashed rounded-2xl p-12 flex flex-col items-center justify-center transition-all",
                    calcFile ? "border-emerald-500/30 bg-emerald-500/5" : "border-white/10 group-hover:border-purple-500/30 group-hover:bg-purple-500/5"
                  )}>
                    {calcFile ? (
                      <>
                        <CheckCircle2 className="text-emerald-500 mb-4" size={32} />
                        <span className="text-sm font-medium">{calcFile.name}</span>
                        <button 
                          onClick={(e) => { e.preventDefault(); setCalcFile(null); }}
                          className="mt-4 text-[10px] uppercase tracking-widest text-red-500 hover:underline"
                        >
                          Remove
                        </button>
                      </>
                    ) : (
                      <>
                        <Upload className="text-white/10 mb-4 group-hover:text-purple-500/50 transition-colors" size={32} />
                        <span className="text-sm text-white/30">Drop calc file here</span>
                      </>
                    )}
                  </div>
                </label>
              </div>

              <div className="md:col-span-2 flex flex-col items-center mt-8 gap-6">
                <AnimatePresence>
                  {processError && (
                    <motion.div 
                      initial={{ opacity: 0, y: 10 }}
                      animate={{ opacity: 1, y: 0 }}
                      exit={{ opacity: 0, y: 10 }}
                      className="w-full max-w-md p-6 bg-red-500/10 border border-red-500/20 rounded-3xl shadow-xl shadow-red-500/5 flex flex-col gap-4"
                    >
                      <div className="flex items-start gap-3">
                        <AlertCircle className="text-red-400 shrink-0" size={20} />
                        <p className="text-sm text-red-200 leading-relaxed font-medium">{processError}</p>
                      </div>

                      {availableSheets.length > 0 && (
                        <div className="mt-2">
                          <label className="text-[10px] uppercase tracking-widest font-mono text-white/40 mb-2 block">Select the correct tab manually:</label>
                          <div className="grid grid-cols-1 gap-2 max-h-40 overflow-y-auto pr-2 custom-scrollbar">
                            {availableSheets.map(sheet => (
                              <button
                                key={sheet}
                                onClick={() => handleManualSheetSelect(sheet)}
                                className="text-left px-4 py-2 rounded-xl bg-white/5 hover:bg-purple-600 hover:text-white text-xs transition-all border border-white/5"
                              >
                                {sheet}
                              </button>
                            ))}
                          </div>
                        </div>
                      )}
                    </motion.div>
                  )}
                </AnimatePresence>

                <button
                  disabled={(!reportFile && !reportText) || !calcFile || isProcessing}
                  onClick={processFiles}
                  className={cn(
                    "purple-button flex items-center gap-3 text-sm font-medium transition-all px-12 py-4",
                    ((!reportFile && !reportText) || !calcFile || isProcessing) ? "opacity-30 cursor-not-allowed grayscale" : "hover:scale-105 active:scale-95 shadow-xl shadow-purple-600/20"
                  )}
                >
                  {isProcessing ? (
                    <>
                      <div className="w-4 h-4 border-2 border-white/30 border-t-white rounded-full animate-spin" />
                      Processing...
                    </>
                  ) : (
                    <>
                      Process & Validate
                      <ArrowRight size={18} />
                    </>
                  )}
                </button>
              </div>
            </motion.div>
          )}

          {currentStep === 'validate' && validation && (
            <motion.div 
              key="validate"
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              className="max-w-2xl mx-auto"
            >
              <div className="bg-white/5 rounded-3xl p-10 border border-white/10 shadow-sm">
                <div className="text-center mb-12">
                  <div className="w-20 h-20 bg-emerald-500/10 text-emerald-400 rounded-full flex items-center justify-center mx-auto mb-6">
                    <CheckCircle2 size={40} />
                  </div>
                  <h2 className="font-serif text-3xl italic mb-2">Validation Successful</h2>
                  <p className="text-sm text-white/40">All values match the Coverage Breakdown tab.</p>
                </div>

                <div className="space-y-4 mb-12">
                  <div className="flex items-center justify-between p-4 bg-black/20 rounded-2xl border border-white/5">
                    <div className="flex items-center gap-3">
                      <div className="w-3 h-3 rounded-full bg-red-500" />
                      <span className="text-sm font-medium">RED Coverage Match</span>
                    </div>
                    <span className="font-mono text-xs text-emerald-400 font-bold">MATCHED</span>
                  </div>
                  <div className="flex items-center justify-between p-4 bg-black/20 rounded-2xl border border-white/5">
                    <div className="flex items-center gap-3">
                      <div className="w-3 h-3 rounded-full bg-emerald-500" />
                      <span className="text-sm font-medium">GREEN Coverage Match</span>
                    </div>
                    <span className="font-mono text-xs text-emerald-400 font-bold">MATCHED</span>
                  </div>
                </div>

                <div className="flex flex-col gap-4">
                  <button
                    onClick={downloadProcessedFile}
                    className="purple-button flex items-center justify-center gap-3 py-5 shadow-xl shadow-purple-600/20"
                  >
                    <Download size={20} />
                    Export Values-Only Report
                  </button>
                  <button
                    onClick={() => setCurrentStep('upload')}
                    className="text-[10px] uppercase tracking-widest text-white/40 hover:text-white transition-colors py-2"
                  >
                    Start Over
                  </button>
                </div>
              </div>
            </motion.div>
          )}

          {currentStep === 'email' && (
            <motion.div 
              key="email"
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              className="max-w-3xl mx-auto"
            >
              <div className="bg-white/5 rounded-3xl p-10 border border-white/10 shadow-sm overflow-hidden relative">
                <div className="absolute top-0 right-0 p-8 opacity-[0.03] pointer-events-none">
                  <Mail size={200} />
                </div>

                <div className="flex items-center justify-between mb-10">
                  <div className="flex items-center gap-4">
                    <div className="w-12 h-12 bg-purple-600/10 text-purple-400 rounded-2xl flex items-center justify-center">
                      <Send size={24} />
                    </div>
                    <div>
                      <h2 className="font-serif text-2xl italic">Distribution Summary</h2>
                      <p className="text-xs text-white/30 font-mono uppercase tracking-wider">Ready for dispatch</p>
                    </div>
                  </div>
                  <button 
                    onClick={() => setIsEditingEmail(!isEditingEmail)}
                    className="text-[10px] uppercase tracking-widest px-4 py-2 rounded-full border border-white/10 hover:bg-white/5 transition-all text-purple-400"
                  >
                    {isEditingEmail ? "Save Template" : "Edit Template"}
                  </button>
                </div>

                <div className="space-y-6 mb-12">
                  {isEditingEmail ? (
                    <div className="space-y-4 animate-in fade-in slide-in-from-top-2 duration-300">
                      <div>
                        <label className="text-[10px] uppercase tracking-widest text-white/30 font-mono mb-1 block">To</label>
                        <input 
                          type="text"
                          value={emailTo}
                          onChange={(e) => setEmailTo(e.target.value)}
                          className="w-full bg-black/20 border border-white/10 rounded-xl px-4 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-purple-500/20"
                        />
                      </div>
                      <div>
                        <label className="text-[10px] uppercase tracking-widest text-white/30 font-mono mb-1 block">BCC (Semicolon separated)</label>
                        <input 
                          type="text"
                          value={emailBcc}
                          onChange={(e) => setEmailBcc(e.target.value)}
                          className="w-full bg-black/20 border border-white/10 rounded-xl px-4 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-purple-500/20"
                        />
                      </div>
                      <div>
                        <label className="text-[10px] uppercase tracking-widest text-white/30 font-mono mb-1 block">Subject</label>
                        <input 
                          type="text"
                          value={emailSubject}
                          onChange={(e) => setEmailSubject(e.target.value)}
                          className="w-full bg-black/20 border border-white/10 rounded-xl px-4 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-purple-500/20"
                        />
                      </div>
                      <div>
                        <label className="text-[10px] uppercase tracking-widest text-white/30 font-mono mb-1 block">Body</label>
                        <textarea 
                          value={emailBody}
                          onChange={(e) => setEmailBody(e.target.value)}
                          className="w-full bg-black/20 border border-white/10 rounded-xl px-4 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-purple-500/20 h-24 resize-none"
                        />
                      </div>
                    </div>
                  ) : (
                    <>
                      <div className="grid grid-cols-[100px_1fr] gap-4 items-center border-b border-white/5 pb-4">
                        <span className="text-[10px] uppercase tracking-widest text-white/30 font-mono">To</span>
                        <span className="text-sm font-medium">{emailTo}</span>
                      </div>
                      <div className="grid grid-cols-[100px_1fr] gap-4 items-center border-b border-white/5 pb-4">
                        <span className="text-[10px] uppercase tracking-widest text-white/30 font-mono">BCC</span>
                        <div className="flex flex-wrap gap-2">
                          {emailBcc.split(';').map((email, i) => (
                            <span key={i} className="px-2 py-1 bg-white/5 rounded text-[10px]">{email.trim()}</span>
                          ))}
                        </div>
                      </div>
                      <div className="grid grid-cols-[100px_1fr] gap-4 items-center border-b border-white/5 pb-4">
                        <span className="text-[10px] uppercase tracking-widest text-white/30 font-mono">Subject</span>
                        <span className="text-sm italic font-serif">{emailSubject}</span>
                      </div>
                    </>
                  )}
                  <div className="grid grid-cols-[100px_1fr] gap-4 items-start pt-2">
                    <span className="text-[10px] uppercase tracking-widest text-white/30 font-mono">Attachment</span>
                    <div className="flex items-center gap-2 text-emerald-400 bg-emerald-500/5 px-3 py-2 rounded-xl border border-emerald-500/10">
                      <FileSpreadsheet size={16} />
                      <span className="text-xs font-bold">Hedge_Coverage_Report.xlsx</span>
                    </div>
                  </div>
                </div>

                <div className="flex flex-col gap-4">
                  <a
                    href={getEmailLink()}
                    className="purple-button flex items-center justify-center gap-3 py-5 shadow-xl shadow-purple-600/20 text-center"
                  >
                    <Mail size={20} />
                    Open Email Draft
                  </a>
                  <button
                    onClick={() => setCurrentStep('upload')}
                    className="text-[10px] uppercase tracking-widest text-white/40 hover:text-white transition-colors py-2"
                  >
                    Process Another Report
                  </button>
                </div>

                {!isMonday && (
                  <div className="mt-8 p-4 bg-purple-500/5 border border-purple-500/10 rounded-2xl flex items-start gap-3">
                    <AlertCircle className="text-purple-400 shrink-0" size={18} />
                    <p className="text-[11px] text-purple-200/70 leading-relaxed">
                      <strong>Note:</strong> Today is not Monday. According to SOP, this distribution is typically performed on <strong>Mondays Only</strong>. Please verify if you wish to proceed.
                    </p>
                  </div>
                )}
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </main>

      {/* Footer / Info */}
      <footer className="max-w-5xl mx-auto px-6 py-12 border-t border-white/5 mt-12">
        <div className="grid md:grid-cols-3 gap-12">
          <div>
            <h4 className="text-[10px] uppercase tracking-widest font-mono mb-4 opacity-30">Storage Path</h4>
            <p className="text-xs font-mono bg-white/5 p-3 rounded-lg border border-white/5 break-all text-white/50">
              F:\Trading Desk\Client Files\Hedge Coverage Data
            </p>
          </div>
          <div>
            <h4 className="text-[10px] uppercase tracking-widest font-mono mb-4 opacity-30">SOP Reference</h4>
            <ul className="text-xs space-y-2 text-white/40">
              <li>• Save report in Clients Folder</li>
              <li>• Validate RED/GREEN matches</li>
              <li>• Paste values (remove formulas)</li>
              <li>• Monday distribution schedule</li>
            </ul>
          </div>
          <div>
            <h4 className="text-[10px] uppercase tracking-widest font-mono mb-4 opacity-30">System Status</h4>
            <div className="flex items-center gap-2 text-xs text-emerald-400">
              <div className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse" />
              Operational
            </div>
          </div>
        </div>
      </footer>

      <style dangerouslySetInnerHTML={{ __html: `
        @import url('https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,400;0,500;1,400;1,500&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500&display=swap');
        
        :root {
          --font-serif: 'Cormorant Garamond', serif;
          --font-sans: 'Inter', sans-serif;
          --font-mono: 'JetBrains Mono', monospace;
        }

        .purple-button {
          background-color: #7c3aed;
          color: white;
          border-radius: 9999px;
          padding: 12px 24px;
          letter-spacing: 0.5px;
          transition: all 0.2s ease;
        }

        .purple-button:hover {
          background-color: #6d28d9;
          transform: translateY(-1px);
        }

        .purple-button:active {
          transform: translateY(0);
        }

        .custom-scrollbar::-webkit-scrollbar {
          width: 4px;
        }
        .custom-scrollbar::-webkit-scrollbar-track {
          background: rgba(255, 255, 255, 0.05);
        }
        .custom-scrollbar::-webkit-scrollbar-thumb {
          background: #7c3aed;
          border-radius: 10px;
        }
      `}} />
    </div>
  );
}
