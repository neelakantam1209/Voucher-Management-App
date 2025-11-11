import React, { useState, useCallback, useEffect } from 'react';
import { FileUpload } from './components/FileUpload';
import { Voucher } from './components/Voucher';
import { Spinner } from './components/Spinner';
import { parseExcelFile } from './services/parserService';
import { extractDataFromImage } from './services/geminiService';
import { downloadAllVouchersAsPDF } from './utils/pdfUtils';
import type { ExtractedRow } from './types';
import { CHECKED_BY_SIGNATURE_B64, APPROVED_SIGNATURE_B64, DEFAULT_RECEIVER_SIGNATURE_B64, DEFAULT_LOGO_B64 } from './constants';
import { ImageUploadControl } from './components/ImageUploadControl';
import { MultiFileUpload } from './components/MultiFileUpload';
import { fileToDataUrl } from './utils/fileUtils';

const App: React.FC = () => {
  const [extractedData, setExtractedData] = useState<ExtractedRow[]>([]);
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [error, setError] = useState<string | null>(null);
  const [fileName, setFileName] = useState<string>('');
  const [currentIndex, setCurrentIndex] = useState<number>(0);
  const [isDownloadingAll, setIsDownloadingAll] = useState<boolean>(false);

  // State to hold the data file before processing
  const [dataFile, setDataFile] = useState<File | null>(null);

  // Customization state
  const [customLogoUrl, setCustomLogoUrl] = useState<string>(DEFAULT_LOGO_B64);
  const [checkedBySigUrl, setCheckedBySigUrl] = useState<string>(CHECKED_BY_SIGNATURE_B64);
  const [approvedBySigUrl, setApprovedBySigUrl] = useState<string>(APPROVED_SIGNATURE_B64);
  const [branchSignatures, setBranchSignatures] = useState<Record<string, string>>({});


  const getReceiverSignatureForVoucher = useCallback((description: string): string => {
    // 1. Normalize fields: trim whitespace and compare case-insensitively.
    const desc = (description || '').trim();

    // 2. Deterministic branch mapping (highest priority)
    // The map defines the logic: [regexToMatch, [possibleFilenamesToLookFor]]
    const signatureMap: [RegExp, string[]][] = [
        // Rule: If Being contains GRA or GRANULE(S), assign Granules signature.
        [/\b(GRANULES|GRANULE|GRA)\b/i, ['GRANULES', 'GRANULE']],
        [/\b(CHEVELLA|CHE)\b/i, ['CHEVELLA']],
        [/\b(BOLLARAM|BOL)\b/i, ['BOLLARAM']],
        [/\b(BHEL|BHE)\b/i, ['BHEL']],
        [/\b(ECIL|ECI)\b/i, ['ECIL']],
        [/\b(MYP)\b/i, ['MYP']],
        [/\b(MKR)\b/i, ['MKR']],
        [/\b(GHM)\b/i, ['GHM']],
    ];

    for (const [regex, signatureKeys] of signatureMap) {
        if (regex.test(desc)) {
            // 3. Exact filename fallback
            // If the regex matches, check for any of the possible signature filenames
            for (const key of signatureKeys) {
                if (branchSignatures[key]) {
                    // 6. Logging & preview
                    console.log(`Voucher Description: "${description}", Matched Rule: "${regex}", Selected Signature: "${key}"`);
                    return branchSignatures[key];
                }
            }
        }
    }

    // 4. Random/Default fallback (lowest priority)
    const signatureUrls = Object.values(branchSignatures);
    if (signatureUrls.length > 0) {
        // If no explicit mapping matches, randomly pick one signature from the uploaded pool.
        const randomIndex = Math.floor(Math.random() * signatureUrls.length);
        const randomSignatureUrl = signatureUrls[randomIndex];
        console.log(`Voucher Description: "${description}", No rule matched. Selected Random Signature.`);
        return randomSignatureUrl;
    }

    // If no uploaded signatures exist, use the configured default receiver signature.
    console.log(`Voucher Description: "${description}", No rule matched and no signatures in pool. Using Default.`);
    return DEFAULT_RECEIVER_SIGNATURE_B64;
  }, [branchSignatures]);

  const processFile = useCallback(async (file: File) => {
    setIsLoading(true);
    setError(null);
    setExtractedData([]);
    setCurrentIndex(0);

    try {
      let data: ExtractedRow[];
      const fileType = file.type;
      const fileExtension = file.name.split('.').pop()?.toLowerCase() || '';

      if (['xlsx', 'xls', 'csv'].includes(fileExtension) || fileType.includes('spreadsheet') || fileType.includes('csv')) {
        data = await parseExcelFile(file);
      } else if (['image/jpeg', 'image/png'].includes(fileType)) {
        data = await extractDataFromImage(file);
      } else {
        throw new Error('Unsupported file type. Please upload an Excel sheet or an image (JPG, PNG).');
      }
      
      setExtractedData(data);
    } catch (err) {
      console.error(err);
      setError(err instanceof Error ? err.message : 'An unknown error occurred.');
      setExtractedData([]); // Ensure data is cleared on error
    } finally {
      setIsLoading(false);
    }
  }, []);
  
  const handleGenerateClick = useCallback(() => {
    if (dataFile) {
        processFile(dataFile);
    }
  }, [dataFile, processFile]);

  const handleDataFileSelect = useCallback((file: File) => {
    setDataFile(file);
    setFileName(file.name);
    setError(null); // Clear previous errors on new selection
  }, []);

  const createUploadHandler = (setter: React.Dispatch<React.SetStateAction<string>>) =>
    useCallback(async (file: File) => {
      try {
        const dataUrl = await fileToDataUrl(file);
        setter(dataUrl);
      } catch (err) {
        console.error("Error converting image to data URL:", err);
        setError("Failed to upload image. Please try another one.");
      }
    }, [setter]);

  const handleLogoUpload = createUploadHandler(setCustomLogoUrl);
  const handleCheckedBySigUpload = createUploadHandler(setCheckedBySigUrl);
  const handleApprovedBySigUpload = createUploadHandler(setApprovedBySigUrl);

  const handleBranchSignaturesUpload = async (files: FileList | null) => {
    if (!files) return;

    const newSignatures: Record<string, string> = {};
    for (const file of Array.from(files)) {
      try {
        const dataUrl = await fileToDataUrl(file);
        // Extract filename without extension, convert to uppercase for consistent matching
        const fileName = file.name.split('.').slice(0, -1).join('.').toUpperCase();
        if (fileName) {
          newSignatures[fileName] = dataUrl;
        }
      } catch (err) {
        console.error("Error converting image to data URL:", err);
        setError(`Failed to upload image ${file.name}. Please try again.`);
      }
    }
    // Merge with existing signatures, allowing override
    setBranchSignatures(prev => ({ ...prev, ...newSignatures }));
  };

  const handleDownloadAll = () => {
    if (extractedData.length > 0) {
      setIsDownloadingAll(true);
    }
  };
  
  useEffect(() => {
    if (isDownloadingAll) {
      const download = async () => {
        await new Promise(resolve => setTimeout(resolve, 100));
        await downloadAllVouchersAsPDF(fileName, 'voucher-dl-');
        setIsDownloadingAll(false);
      };
      download();
    }
  }, [isDownloadingAll, fileName]);

  const handlePrevious = () => {
    setCurrentIndex(prev => Math.max(0, prev - 1));
  };

  const handleNext = () => {
    setCurrentIndex(prev => Math.min(extractedData.length - 1, prev + 1));
  };

  const handleStartOver = () => {
      setExtractedData([]);
      setDataFile(null);
      setError(null);
      setFileName('');
      setCurrentIndex(0);
      setIsLoading(false);
  };

  const NavButton: React.FC<{ onClick: () => void; disabled: boolean; children: React.ReactNode }> = ({ onClick, disabled, children }) => (
    <button
      onClick={onClick}
      disabled={disabled}
      className="px-4 py-2 bg-gray-700 text-white font-bold rounded-lg transition-colors hover:bg-gray-600 disabled:bg-gray-800 disabled:text-gray-500 disabled:cursor-not-allowed"
    >
      {children}
    </button>
  );

  const renderInitialScreen = () => (
    <div className="space-y-8">
        <div className="bg-gray-800/50 backdrop-blur-sm rounded-xl shadow-2xl p-6 md:p-8 border border-gray-700">
            <h2 className="text-xl font-semibold text-gray-200 mb-2">Step 1: Customize Your Vouchers <span className="text-gray-400 font-normal">(Optional)</span></h2>
            <p className="text-gray-400 mb-6 text-sm">Upload a custom logo and signatures. If left blank, defaults will be used.</p>
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-6">
                 <ImageUploadControl
                    id="logo-upload"
                    label="Custom Logo"
                    previewUrl={customLogoUrl}
                    defaultUrl={DEFAULT_LOGO_B64}
                    onFileSelect={handleLogoUpload}
                    disabled={isLoading}
                  />
                  <ImageUploadControl
                    id="checked-by-upload"
                    label="'Checked By' Sig"
                    previewUrl={checkedBySigUrl}
                    defaultUrl={CHECKED_BY_SIGNATURE_B64}
                    onFileSelect={handleCheckedBySigUpload}
                    disabled={isLoading}
                  />
                  <ImageUploadControl
                    id="approved-by-upload"
                    label="'Approved By' Sig"
                    previewUrl={approvedBySigUrl}
                    defaultUrl={APPROVED_SIGNATURE_B64}
                    onFileSelect={handleApprovedBySigUpload}
                    disabled={isLoading}
                  />
            </div>
             <div className="mt-8 pt-6 border-t border-gray-700">
                <h3 className="text-lg font-semibold text-gray-200 mb-2">Upload Receiver Signatures</h3>
                <p className="text-gray-400 mb-4 text-sm">
                    Upload all branch signatures at once. The signature used will be automatically matched based on the voucher's "Being" description.
                    <br />
                    <strong>Important:</strong> Name files like <code className="bg-gray-900 text-gray-300 px-1 rounded-md text-xs">GRANULES.png</code>, <code className="bg-gray-900 text-gray-300 px-1 rounded-md text-xs">BHEL.jpg</code>, or <code className="bg-gray-900 text-gray-300 px-1 rounded-md text-xs">MYP.png</code>.
                </p>
                <MultiFileUpload onFileSelect={handleBranchSignaturesUpload} disabled={isLoading} />
                {Object.keys(branchSignatures).length > 0 && (
                    <div className="mt-4 p-4 bg-gray-900/50 rounded-lg">
                        <p className="text-sm font-semibold text-gray-300">Uploaded Signatures:</p>
                        <ul className="list-disc list-inside text-sm text-gray-400 mt-2 grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-4 gap-x-4 gap-y-1">
                            {Object.keys(branchSignatures).map(name => 
                                <li key={name} className="truncate" title={name}>
                                    <span className="text-green-400 mr-1">✓</span>{name}
                                </li>
                            )}
                        </ul>
                    </div>
                )}
            </div>
        </div>

        <div className="bg-gray-800/50 backdrop-blur-sm rounded-xl shadow-2xl p-6 md:p-8 border border-gray-700">
          <h2 className="text-xl font-semibold text-gray-200 mb-4">Step 2: Upload Your Data File</h2>
          <FileUpload onFileSelect={handleDataFileSelect} disabled={isLoading} />
          {fileName && (
            <div className="mt-4 text-center">
              <p className="text-sm text-gray-300">Selected file: <span className="font-semibold text-blue-400">{fileName}</span></p>
            </div>
          )}
        </div>

        <div className="flex justify-center">
            <button
                onClick={handleGenerateClick}
                disabled={!dataFile || isLoading}
                className="w-full md:w-auto text-lg font-bold text-white bg-blue-600 rounded-lg px-8 py-3 transition-all transform hover:bg-blue-500 hover:scale-105 disabled:bg-gray-700 disabled:cursor-not-allowed disabled:scale-100 flex items-center justify-center"
            >
                {isLoading ? <><Spinner /> Generating...</> : 'Generate Vouchers'}
            </button>
        </div>
    </div>
  );

  const renderVoucherScreen = () => (
    <div className="w-full max-w-5xl mx-auto">
        <div className="flex justify-between items-center mb-6 flex-wrap gap-4">
            <h2 className="text-xl md:text-2xl font-bold text-white">Generated Vouchers</h2>
            <div className="flex items-center space-x-2">
                <NavButton onClick={handlePrevious} disabled={currentIndex === 0}>Previous</NavButton>
                <span className="text-gray-300 font-semibold">{currentIndex + 1} / {extractedData.length}</span>
                <NavButton onClick={handleNext} disabled={currentIndex === extractedData.length - 1}>Next</NavButton>
            </div>
        </div>

        {isDownloadingAll && extractedData.map((data, index) => (
            <div key={index} className="absolute -left-[9999px] top-0">
                <Voucher 
                    id={`voucher-dl-${index}`}
                    data={data}
                    receiverSignatureUrl={getReceiverSignatureForVoucher(data.description)}
                    checkedBySigUrl={checkedBySigUrl}
                    approvedBySigUrl={approvedBySigUrl}
                    logoUrl={customLogoUrl}
                />
            </div>
        ))}
        
        <Voucher
            id="voucher-display"
            data={extractedData[currentIndex]}
            receiverSignatureUrl={getReceiverSignatureForVoucher(extractedData[currentIndex].description)}
            checkedBySigUrl={checkedBySigUrl}
            approvedBySigUrl={approvedBySigUrl}
            logoUrl={customLogoUrl}
        />

        <div className="mt-8 flex justify-center flex-wrap gap-4">
            <button
                onClick={handleDownloadAll}
                className="px-6 py-3 bg-green-600 text-white font-bold rounded-lg transition-colors hover:bg-green-500 disabled:bg-gray-700 disabled:cursor-not-allowed"
                disabled={isDownloadingAll}
            >
                {isDownloadingAll ? 'Downloading...' : 'Download All as PDF'}
            </button>
            <button
                onClick={handleStartOver}
                className="px-6 py-3 bg-gray-700 text-white font-bold rounded-lg transition-colors hover:bg-gray-600"
            >
                Start Over
            </button>
        </div>
    </div>
  );

  return (
    <main className="container mx-auto px-4 py-8 md:py-12">
        <header className="text-center mb-10">
            <h1 className="text-3xl md:text-4xl font-extrabold text-transparent bg-clip-text bg-gradient-to-r from-blue-400 to-teal-300">
                Advanced Voucher Generator
            </h1>
            <p className="mt-2 text-gray-400 max-w-2xl mx-auto">
                Automatically generate professional payment vouchers from your Excel files or images.
            </p>
        </header>

        {error && (
          <div className="my-6 p-4 bg-red-900/50 text-red-300 border border-red-700 rounded-lg text-center">
            <p className="font-bold">An Error Occurred</p>
            <p className="text-sm">{error}</p>
          </div>
        )}

        {extractedData.length > 0 ? renderVoucherScreen() : renderInitialScreen()}

    </main>
  );
};

export default App;