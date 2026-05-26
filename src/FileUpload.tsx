import { useRef, useState } from "react";
import './FileUpload.css'

interface uploadProps {
  handleFile: (file: File) => void;
}

export default function FileUpload({handleFile}: uploadProps) {
  const [isDragging, setIsDragging] = useState(false);

  const inputRef = useRef<HTMLInputElement | null>(null);

  const handleFileUpload = (
    e: React.ChangeEvent<HTMLInputElement>
  ) => {
    const files = e.target.files;

    if (!files || files.length === 0) return;

    handleFile(files[0]);
  };

  const handleDrop = (
    e: React.DragEvent<HTMLDivElement>
  ) => {
    e.preventDefault();
    setIsDragging(false);

    const files = e.dataTransfer.files;

    if (!files || files.length === 0) return;

    handleFile(files[0]);
  };

  return (
    <div
      className={`input-box ${isDragging ? "dragging" : ""}`}
      onDragOver={(e) => {
        e.preventDefault();
        setIsDragging(true);
      }}
      onDragLeave={(e) => { 
        e.preventDefault();
        setIsDragging(false);
      }}
      onDrop={handleDrop}
      onClick={() => inputRef.current?.click()}
    >
      <input
        ref={inputRef}
        type="file"
        accept=".xlsx,.xls"
        onChange={handleFileUpload}
        hidden
      />

      <p>
        Drag & drop Excel file here
        <br />
        or click to upload
      </p>
    </div>
  );
}