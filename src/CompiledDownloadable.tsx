import { useCallback } from "react";
import ExcelJS from "exceljs";
import './App.css'
import { writeFile } from "@tauri-apps/plugin-fs";
import { save } from "@tauri-apps/plugin-dialog";

interface CompiledDownloadableProps {
    sickLeaveData : leaveObject[][],
    vacationLeaveData : leaveObject[][],
}

function CompiledDownloadable({sickLeaveData, vacationLeaveData} : CompiledDownloadableProps){

    const createLeaveSheet = useCallback(async(leaveType : number)=>{

        //getting the template
        const response = await fetch("/templates/LEAVE-TEMPLATE.xlsx");
        const arrayBuffer = await response.arrayBuffer();
        
        //creating the new workbook for download
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(arrayBuffer);
        
        //copying the Blank Format sheet from the tempalte
        const templateSheet = workbook.getWorksheet("Blank Format");

        if(!templateSheet) return;

        if (leaveType == 0) {
            sickLeaveData.forEach((item)=>{
                
                //sheet name
                const dynamicSheetName = `${item[0].__EMPTY_1}`

                //creating new sheet on the workbook for download,
                //applying the setup of templateSheet
                const newSheet = workbook.addWorksheet(dynamicSheetName, {
                    pageSetup: templateSheet.pageSetup, 
                    views: templateSheet.views  
                });

                // Copy column widths
                templateSheet.columns.forEach((col, index) => {
                    if (col.width) {
                        newSheet.getColumn(index + 1).width = col.width;
                    }
                });

                // Copy rows (styles + values) — no merge checks needed here
                templateSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
                    const newRow = newSheet.getRow(rowNumber);
                    if (row.height) newRow.height = row.height;

                    for (let colNumber = 1; colNumber <= templateSheet.columnCount; colNumber++) {
                        const cell = row.getCell(colNumber);
                        const newCell = newRow.getCell(colNumber);

                        newCell.style = JSON.parse(JSON.stringify(cell.style || {}));

                        if (cell.value !== null && cell.value !== undefined) {
                            newCell.value = cell.value;
                        }
                    }
                    newRow.commit();
                });

                // Copy merges
                const merges: string[] = (templateSheet as any).model?.merges ?? [];
                merges.forEach((mergeRange: string) => {
                    newSheet.mergeCells(mergeRange);
                });

                //Name, ID, and Leave Type on the header       
                newSheet.getCell("C2").value = `${item[0].__EMPTY_1}`;
                newSheet.getCell("C4").value = `${item[0].__EMPTY}`;
                newSheet.getCell("C5").value = "Sick Leave";

                //11 is First Row after all the header and blot
                let i = 11
                item.forEach((element : leaveObject) => {
                    // sheet.getCell(`B${i}`).value = new Date(element.__EMPTY_2).toISOString().split("T")[0];
                    newSheet.getCell(`B${i}`).value = new Date(element.__EMPTY_2);
                    i++
                })
            })            
        } 
        
        else if (leaveType == 1){
            vacationLeaveData.forEach((item)=>{
                
                //sheet name
                const dynamicSheetName = `${item[0].__EMPTY_1}`

                //creating new sheet on the workbook for download,
                //applying the setup of templateSheet
                const newSheet = workbook.addWorksheet(dynamicSheetName, {
                    pageSetup: templateSheet.pageSetup, 
                    views: templateSheet.views  
                });

                // Copy column widths
                templateSheet.columns.forEach((col, index) => {
                    if (col.width) {
                        newSheet.getColumn(index + 1).width = col.width;
                    }
                });

                // Copy rows (styles + values) — no merge checks needed here
                templateSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
                    const newRow = newSheet.getRow(rowNumber);
                    if (row.height) newRow.height = row.height;

                    for (let colNumber = 1; colNumber <= templateSheet.columnCount; colNumber++) {
                        const cell = row.getCell(colNumber);
                        const newCell = newRow.getCell(colNumber);

                        newCell.style = JSON.parse(JSON.stringify(cell.style || {}));

                        if (cell.value !== null && cell.value !== undefined) {
                            newCell.value = cell.value;
                        }
                    }
                    newRow.commit();
                });

                // Copy merges
                const merges: string[] = (templateSheet as any).model?.merges ?? [];
                merges.forEach((mergeRange: string) => {
                    newSheet.mergeCells(mergeRange);
                });

                //Name, ID, and Leave Type on the header       
                newSheet.getCell("C2").value = `${item[0].__EMPTY_1}`;
                newSheet.getCell("C4").value = `${item[0].__EMPTY}`;
                newSheet.getCell("C5").value = "Vacation Leave";

                //11 is First Row after all the header and blot
                let i = 11
                item.forEach((element : leaveObject) => {
                    // sheet.getCell(`B${i}`).value = new Date(element.__EMPTY_2).toISOString().split("T")[0];
                    newSheet.getCell(`B${i}`).value = new Date(element.__EMPTY_2);
                    i++
                })

                
            })

        }



    
        async function downloadExcel(workbook: ExcelJS.Workbook) {
            const buffer = await workbook.xlsx.writeBuffer();
            const uint8Array = new Uint8Array(buffer as ArrayBuffer);

            // Detect Tauri environment
            const isTauri = "__TAURI_INTERNALS__" in window;

            let month : Date | null    

            if (leaveType == 0){
                month = (new Date(sickLeaveData[0][0].__EMPTY_2))
                // console.log("month set to: ", new Date(sickLeaveData[0][0].__EMPTY_2))
            } else if (leaveType == 1){
                month = (new Date(vacationLeaveData[0][0].__EMPTY_2))
            } else month = null

            if(!month) return

            if (isTauri) {
                // Opens native OS "Save As" dialog — user picks location + filename
                const filePath = await save({
                    defaultPath: `${month.toLocaleDateString('en-US', { month: 'long' })} ${leaveType == 0 ? "SL" : "VL"}.xlsx`,
                    filters: [
                        {
                        name: "Excel Spreadsheet",
                        extensions: ["xlsx"],
                        },
                    ],
                    });

                if (!filePath) return; // User cancelled — no ghost downloads

                await writeFile(filePath, uint8Array);
            } else {
                // Original browser flow unchanged
                const blob = new Blob([buffer], {
                type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    });
                    const url = URL.createObjectURL(blob);
                    const a = document.createElement("a");
                    a.href = url;
                    a.download = `${month.toLocaleDateString('en-US', { month: 'long' })} ${leaveType == 0 ? "SL" : "VL"}.xlsx`;
                    a.click();
                    URL.revokeObjectURL(url);
                }
            }

            downloadExcel(workbook)
        }, [sickLeaveData,vacationLeaveData])




    return(
        <>        
            {sickLeaveData.length !== 0 ?            
            <div 
                onClick={async () => {
                    await createLeaveSheet(0)
                }}
                className="download-card"
                >
                <h4>    
                    {sickLeaveData[0][0].__EMPTY_2.toLocaleDateString('en-US', { month: 'long' })} Sick Leave     
                </h4>

                
                {/* {`${userData[0].__EMPTY_1}.xlsx`} */}
            </div> : ""}

            {vacationLeaveData.length !== 0 ?             
            <div 
                onClick={async () => {
                    await createLeaveSheet(1)
                }}
                className="download-card"
                >
                <h4>
                    {vacationLeaveData[0][0].__EMPTY_2.toLocaleDateString('en-US', { month: 'long' })} Vacation Leave               
                </h4>

                 
                {/* {`${userData[0].__EMPTY_1}.xlsx`} */}
            </div> : ""}


        </>

    )
}

export default CompiledDownloadable

type leaveObject = {
  __EMPTY : string;
  __EMPTY_1 : string;
  __EMPTY_2 : Date;
  __EMPTY_3 : Date;
  __EMPTY_5 : number;
  __EMPTY_6 : string
}