import { useEffect, useState } from "react";
import * as XLSX from "xlsx"
import './App.css'
import Downloadable from "./Downloadable";
import FileUpload from "./FileUpload";


// type leaveObject = Record<string, unknown>;

function App() {

  //holds the data of sickleave and vacation leave
  //arrays of arrays, the first array separated per person
  //the inner arrays contain's  the person's leave schedules
  const [sickLeaveData, setSickLeaveData] = useState<leaveObject[][] | null>(null)
  const [vacationLeaveData, setVacationLeaveData] = useState<leaveObject[][] | null>(null)

  //triggers as the sickLeaveData and vacationData changes
  //will be used to display downloadables
  useEffect(()=>{
    console.log(sickLeaveData)
    console.log(vacationLeaveData)
  },[sickLeaveData, vacationLeaveData])

  //triggers when the input change when it recieves a file
  // const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    
  //   //puts said file into constant named files
  //   const file = e.target.files?.[0];

  //   //checks constant file is empty as the
  //   // function may run on load resulting in error
  //   if (!file) return;

  //   //reader from sheetJS
  //   const reader = new FileReader();

  //   reader.onload = (event) => {
  //     const result = event.target?.result;

  //     //if somehow the result is empty
  //     if (!result) return;

  //     //declares how sheetJS should read the file
  //     //celldates true since it contains preformatted cells
  //     const workbook = XLSX.read(result, { type: "binary", cellDates : true, cellStyles : true});

  //     const sheetName = workbook.SheetNames[0];
  //     const worksheet = workbook.Sheets[sheetName];

  //     //converts the sheet to a JSON object
  //     const jsonData: ExcelRow[] = XLSX.utils.sheet_to_json(worksheet);

  //     //removes the bloat and only shows data
  //     const extractedData = extractData(jsonData)

  //     const tempSickLeaveData = processSickLeaveData(extractedData)
  //     const tempVacationLeaveData = processVacationLeaveData(extractedData)

  //     setSickLeaveData(tempSickLeaveData)
  //     setVacationLeaveData(tempVacationLeaveData)

  //     // createLeaveSheet(sickLeaveData[1], "Sick Leave")
  //   };
  
  //   reader.readAsBinaryString(file);

  // };

  const handleFile = (file: File)=>{
    //checks constant file is empty as the
    // function may run on load resulting in error
    if (!file) return;

    //reader from sheetJS
    const reader = new FileReader();

    reader.onload = (event) => {
      const result = event.target?.result;

      //if somehow the result is empty
      if (!result) return;

      //declares how sheetJS should read the file
      //celldates true since it contains preformatted cells
      const workbook = XLSX.read(result, { type: "binary", cellDates : true, cellStyles : true});

      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];

      //converts the sheet to a JSON object
      const jsonData: leaveObject[] = XLSX.utils.sheet_to_json(worksheet);

      //removes the bloat and only shows data
      const extractedData = extractData(jsonData)

      // console.log(extractedData)

      // processSickLeaveData(extractedData)
      const tempSickLeaveData = processSickLeaveData(extractedData)
      const tempVacationLeaveData = processVacationLeaveData(extractedData)

      // console.log(tempSickLeaveData)
      // console.log(tempVacationLeaveData)

      setSickLeaveData(tempSickLeaveData)
      setVacationLeaveData(tempVacationLeaveData)

      // createLeaveSheet(sickLeaveData[1], "Sick Leave")
    };
  
    reader.readAsBinaryString(file);
  }

  //extracts only the leave rows
  //separates sick and vacation leave 0, 1 respectively
  const extractData = (jsonData : leaveObject[]) => {
    // console.log(jsonData[7].__EMPTY)
    let currentGroup : leaveObject[] = []
    const leaveGroupings : leaveObject[][] = []

    jsonData.forEach( num => {
      if(isNaN(Number(num.__EMPTY))){
        if(currentGroup.length !== 0){
          leaveGroupings.push(currentGroup)
          currentGroup = []
        }

        return
      }

      // console.log(num)
      currentGroup.push(num)

    })

    return leaveGroupings
  }

  //takes sick leave only from the extracted data
  //returns an array of array
  //each element on the first array is the employee
  //each employee has their own array of leaves
  const processSickLeaveData = (extractedData : leaveObject[][]) => {
      //takes sickleave data only 
      const groupedSickLeaveData = Object.groupBy(extractedData[0], item => item.__EMPTY)
      // console.log(groupedSickLeaveData)

      // //sorts the sick leave data array by their dates
      const groupedSickLeaveDataArray = Object.values(groupedSickLeaveData).filter((element): element is leaveObject[] => element !== undefined);
      groupedSickLeaveDataArray.forEach(element => {
        if(!element) return
        element.sort((a : leaveObject, b : leaveObject) => a.__EMPTY_2.getTime() - b.__EMPTY_2.getTime())
      });
      return groupedSickLeaveDataArray
  }
  const processVacationLeaveData = (extractedData : leaveObject[][]) => {
  //     //takes vacation leave data only
      const groupedVacationLeaveData = Object.groupBy(extractedData[1], item => item.__EMPTY)

  //     //sorts the sick leave data array by their dates
      const groupedVacationLeaveDataArray = Object.values(groupedVacationLeaveData).filter((element): element is leaveObject[] => element !== undefined);
      groupedVacationLeaveDataArray.forEach(element => {
        if(!element) return
            element.sort((a : leaveObject, b : leaveObject) => a.__EMPTY_2.getTime() - b.__EMPTY_2.getTime())
      });

      return groupedVacationLeaveDataArray
  }



  return (
    <>

      <section id="center">
        <div style={{ padding: 20, width : "100%" }}>
          <h1>Excel Editor</h1>
          
          <div className="box-holder">
            <FileUpload handleFile={handleFile}/>
            {/* <div className="input-box">
              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={handleFileUpload}
              />              
            </div> */}
            <div className="output-box"> 
              {sickLeaveData?.map((item, index)=><Downloadable userData={item} key={index} leaveType={"Sick Leave"}/>)}
              {vacationLeaveData?.map((item, index)=><Downloadable userData={item} key={index} leaveType={"Vacation Leave"}/>)}
              {sickLeaveData?.map((item, index)=><Downloadable userData={item} key={index} leaveType={"Sick Leave"}/>)}
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
