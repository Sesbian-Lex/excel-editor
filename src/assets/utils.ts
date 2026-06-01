import { type Dispatch, type SetStateAction } from 'react';
import * as XLSX from "xlsx"

interface leaveStateProps {
  value: leaveObject[][] | null;
  // This is the accurate type for a useState setter function
  setValue: Dispatch<SetStateAction<leaveObject[][] | null>>; 
}

interface fileStateProps {
  value: File | null ;
  // This is the accurate type for a useState setter function
  setValue: Dispatch<SetStateAction<File | null>>; 
}

  const handleFile = (file: File, sickLeave : leaveStateProps, vacationLeave : leaveStateProps, fileState : fileStateProps)=>{
    //checks constant file is empty as the
    // function may run on load resulting in error
    if (!file) return;

    fileState.setValue(file)

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

      sickLeave.setValue(tempSickLeaveData)
      vacationLeave.setValue(tempVacationLeaveData)

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

  export {handleFile}

  type leaveObject = {
  __EMPTY : string;
  __EMPTY_1 : string;
  __EMPTY_2 : Date;
  __EMPTY_3 : Date;
  __EMPTY_5 : number;
  __EMPTY_6 : string
}
