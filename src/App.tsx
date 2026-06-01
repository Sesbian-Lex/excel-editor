import { useEffect, useState } from "react";
// import * as XLSX from "xlsx"
import './App.css'
import Downloadable from "./Downloadable";
import FileUpload from "./FileUpload";
import { readFile } from '@tauri-apps/plugin-fs';
import CompiledDownloadable from "./CompiledDownloadable";

import { handleFile } from "./assets/utils";

// type leaveObject = Record<string, unknown>;

function App() {

  //holds the data of sickleave and vacation leave
  //arrays of arrays, the first array separated per person
  //the inner arrays contain's  the person's leave schedules
  const [sickLeaveData, setSickLeaveData] = useState<leaveObject[][] | null>(null)
  const [vacationLeaveData, setVacationLeaveData] = useState<leaveObject[][] | null>(null)

  //holds the uploaded file
  const [file, setFile] = useState<File | null>(null)

  //triggers as the sickLeaveData and vacationData changes
  //will be used to display downloadables
  useEffect(()=>{
    console.log(sickLeaveData)
    console.log(vacationLeaveData)
  },[sickLeaveData, vacationLeaveData])


  async function processFile(input: File | string) {
  if (typeof input === "string") {
    // TAURI PATH → convert to File
    const fileUint8Array = await readFile(input); 
    
    // 2. Extract the file name from the path (e.g., "JANUARY.xls")
    const fileName = input.split('/').pop() || 'uploaded_file.xls';

    // 3. Convert the Uint8Array into a standard browser File object
    // Note: 'application/vnd.ms-excel' is the standard MIME type for .xls
    const fileObject = new File([fileUint8Array], fileName, {
      type: 'application/vnd.ms-excel', 
    });

    handleFile(
      fileObject, 
      {
        value : sickLeaveData,
        setValue : setSickLeaveData
      },
      {
        value : vacationLeaveData,
        setValue : setVacationLeaveData
      },
      {
        value : file,
        setValue : setFile
      }
    ) 
  } else {
    // BROWSER FILE
        handleFile(
      input, 
      {
        value : sickLeaveData,
        setValue : setSickLeaveData
      },
      {
        value : vacationLeaveData,
        setValue : setVacationLeaveData
      },
      {
        value : file,
        setValue : setFile
      }
    ) 
  }
}

  const removeFile = () => {
    setFile(null)
    setSickLeaveData(null)
    setVacationLeaveData(null)
  }


  return (
    <>

      <section id="center">
        <div style={{ padding: 20, width : "100%" }}>
          <h1>Excel Editor</h1>
          
          <div className="box-holder">
            {file ? <div className="input-box">
              <span style={{ textOverflow: 'ellipsis', overflow: 'hidden', whiteSpace: 'nowrap' }}>
                📄 {file.name} ({(file.size / 1024).toFixed(1)} KB)
              </span>
              <button 
                onClick={removeFile} 
                className="remove-button"
              >
                X
              </button>
            </div>
            
            : <FileUpload processFile={processFile}/>}
            
            <div className="output-box"> 
              <h2 className="grid-title">Compiled Leave Data</h2>
              <div className="grid-box">
                {!sickLeaveData ? "" : 
                  !vacationLeaveData ? "" :
                    <CompiledDownloadable 
                      sickLeaveData={sickLeaveData} 
                      vacationLeaveData={vacationLeaveData}
                    />
                }
              </div>
        
              <br/>

              <h2 className="grid-title">Sick Leave Data</h2>              
              <div className="grid-box">
                {sickLeaveData?.map((item, index)=><Downloadable userData={item} key={index} leaveType={"Sick Leave"}/>)}
              </div>

                <br/>

              <h2 className="grid-title">Vacation Leave Data</h2>
              <div className="grid-box">
                {vacationLeaveData?.map((item, index)=><Downloadable userData={item} key={index} leaveType={"Vacation Leave"}/>)}
              </div>




            </div>         
          </div>


          {/* <pre>{JSON.stringify(data, null, 2)}</pre> */}
        </div>
      </section>


      {/* <div className="ticks"></div> */}
    </>
  )
}

export default App

type leaveObject = {
  __EMPTY : string;
  __EMPTY_1 : string;
  __EMPTY_2 : Date;
  __EMPTY_3 : Date;
  __EMPTY_5 : number;
  __EMPTY_6 : string
}
