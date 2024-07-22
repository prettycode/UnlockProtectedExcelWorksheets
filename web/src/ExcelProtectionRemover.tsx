import React, { useState, useRef, ChangeEvent } from "react";

interface Worksheet {
  name: string;
  protected: boolean;
}

const ExcelProtectionRemover: React.FC = () => {
  const [file, setFile] = useState<File | null>(null);
  const [buffer, setBuffer] = useState<ArrayBuffer | null>(null);
  const [worksheets, setWorksheets] = useState<Worksheet[]>([]);
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileChange = (event: ChangeEvent<HTMLInputElement>) => {
    const selectedFile = event.target.files?.[0];
    if (selectedFile && selectedFile.name.endsWith(".xlsx")) {
      setIsLoading(true);
      setFile(selectedFile);
      const reader = new FileReader();
      reader.onload = (e: ProgressEvent<FileReader>) => {
        const result = e.target?.result;
        if (result instanceof ArrayBuffer) {
          setBuffer(result);
          setWorksheets(getDummyWorksheets());
          setIsLoading(false);
        }
      };
      reader.readAsArrayBuffer(selectedFile);
    } else {
      alert("Please select a valid .xlsx file");
      setFile(null);
      setBuffer(null);
      setWorksheets([]);
      if (fileInputRef.current) {
        fileInputRef.current.value = "";
      }
    }
  };

  const getDummyWorksheets = (): Worksheet[] => {
    return ["Sheet1", "Sheet2", "Sheet3", "Sheet4"].map((name) => ({
      name,
      protected: Math.random() > 0.5,
    }));
  };

  const handleRemoveProtection = () => {
    if (file && buffer) {
      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = file.name;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }
  };

  const hasProtectedWorksheets = worksheets.some((sheet) => sheet.protected);

  return (
    <div style={styles.container}>
      <div style={styles.fileInputContainer}>
        <label htmlFor="file-upload" style={styles.fileInputLabel}>
          <svg
            style={styles.uploadIcon}
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="2"
            strokeLinecap="round"
            strokeLinejoin="round"
          >
            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
            <polyline points="17 8 12 3 7 8" />
            <line x1="12" y1="3" x2="12" y2="15" />
          </svg>
          <span style={styles.fileInputText}>
            {file ? file.name : "Choose Excel File"}
          </span>
        </label>
        <input
          id="file-upload"
          type="file"
          accept=".xlsx"
          onChange={handleFileChange}
          style={styles.hiddenFileInput}
          ref={fileInputRef}
        />
      </div>

      {isLoading && <p style={styles.loadingText}>Loading file contents...</p>}

      {!isLoading && file && (
        <>
          <div style={styles.worksheetList}>
            <h3 style={styles.heading}>Worksheets in {file.name}:</h3>
            {worksheets.map((sheet, index) => (
              <div key={index} style={styles.worksheetItem}>
                <span>{sheet.name}</span>
                <span
                  style={
                    sheet.protected
                      ? styles.protectedIcon
                      : styles.unprotectedIcon
                  }
                >
                  {sheet.protected ? "🔒" : "🔓"}
                </span>
              </div>
            ))}
          </div>

          <div style={styles.buttonContainer}>
            <button
              onClick={handleRemoveProtection}
              disabled={!hasProtectedWorksheets}
              style={
                hasProtectedWorksheets ? styles.button : styles.disabledButton
              }
              title={
                !hasProtectedWorksheets
                  ? "No protected worksheets"
                  : "Remove protection from worksheets"
              }
            >
              Remove Protection
            </button>
          </div>

          {!hasProtectedWorksheets && (
            <div style={styles.alert}>
              <p>There are no protected worksheets in this file.</p>
            </div>
          )}
        </>
      )}
    </div>
  );
};

const styles: { [key: string]: React.CSSProperties } = {
  container: {
    fontFamily:
      '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif',
    padding: "20px",
    maxWidth: "500px",
    margin: "0 auto",
    backgroundColor: "#ffffff",
    borderRadius: "12px",
    boxShadow: "0 4px 6px rgba(0, 0, 0, 0.1)",
  },
  fileInputContainer: {
    marginBottom: "20px",
  },
  fileInputLabel: {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    padding: "12px",
    backgroundColor: "#f0f0f0",
    borderRadius: "8px",
    cursor: "pointer",
    transition: "background-color 0.3s",
  },
  uploadIcon: {
    width: "24px",
    height: "24px",
    marginRight: "10px",
    color: "#4a4a4a",
  },
  fileInputText: {
    color: "#4a4a4a",
    fontSize: "16px",
  },
  hiddenFileInput: {
    display: "none",
  },
  worksheetList: {
    marginBottom: "20px",
  },
  heading: {
    fontSize: "18px",
    fontWeight: "600",
    marginBottom: "10px",
    color: "#333",
  },
  worksheetItem: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    padding: "10px 0",
    borderBottom: "1px solid #eaeaea",
  },
  protectedIcon: {
    color: "#e74c3c",
  },
  unprotectedIcon: {
    color: "#2ecc71",
  },
  buttonContainer: {
    display: "flex",
    justifyContent: "center",
    marginBottom: "20px",
  },
  button: {
    backgroundColor: "#3498db",
    border: "none",
    color: "white",
    padding: "12px 24px",
    textAlign: "center",
    textDecoration: "none",
    display: "inline-block",
    fontSize: "16px",
    margin: "4px 2px",
    cursor: "pointer",
    borderRadius: "8px",
    transition: "background-color 0.3s",
  },
  disabledButton: {
    backgroundColor: "#bdc3c7",
    color: "#7f8c8d",
    cursor: "not-allowed",
    padding: "12px 24px",
    textAlign: "center",
    textDecoration: "none",
    display: "inline-block",
    fontSize: "16px",
    margin: "4px 2px",
    borderRadius: "8px",
    border: "none",
  },
  alert: {
    backgroundColor: "#fdf2f2",
    color: "#9b1c1c",
    padding: "12px",
    borderRadius: "8px",
    marginBottom: "20px",
    textAlign: "center",
  },
  loadingText: {
    textAlign: "center",
    color: "#7f8c8d",
  },
};

export default ExcelProtectionRemover;
