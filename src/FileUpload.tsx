import { useRef, useState,useEffect } from "react";
import './FileUpload.css'
import { getCurrentWindow } from "@tauri-apps/api/window";
import { isTauri } from "@tauri-apps/api/core";

interface uploadProps {
  processFile: (input: File | string) => void;
}

export default function FileUpload({processFile}: uploadProps) {
  const [isDragging, setIsDragging] = useState(false);

  const inputRef = useRef<HTMLInputElement | null>(null);

  useEffect(() => {
    // const isTauri =
    //   typeof window !== "undefined" &&
    //   "__TAURI__" in window;

    if (!isTauri()) {
      console.log("running in browser");
      return;
    }

    console.log("running in tauri");

    let unlisten: any;

    const setup = async () => {
      console.log("setup started")

      unlisten = await getCurrentWindow().onDragDropEvent((event) => {
        console.log("TAURI EVENT:", event.payload);
         console.log("RAW EVENT:", JSON.stringify(event));

        if (event.payload.type === "drop") {
          setIsDragging(false);

          const paths = event.payload.paths;
          console.log("paths:", paths);

          if (paths.length > 0) {
            processFile(paths[0]);
          }
        }

        if (event.payload.type === "enter") setIsDragging(true);
        if (event.payload.type === "leave") setIsDragging(false);
      });
    };

    setup();

    return () => unlisten?.();
  }, [processFile]);

  const handleFileUpload = (
    e: React.ChangeEvent<HTMLInputElement>
  ) => {
    const files = e.target.files;

    if (!files || files.length === 0) return;

    processFile(files[0]);
  };

  const handleDrop = (
    e: React.DragEvent<HTMLDivElement>
  ) => {
    e.preventDefault();
    console.log("DROP EVENT FIRED");

    setIsDragging(false);

    const files = e.dataTransfer.files;
      console.log(files)

    if (!files || files.length === 0) return;

    processFile(files[0]);
  };

  return (
    <div
      className={`input-box ${isDragging ? "dragging" : ""}`}
        onDragOver={(e) => {
          e.preventDefault();

          if (!isDragging) {
            setIsDragging(true);
          }
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