import { useCallback } from "react";
import ExcelJS from "exceljs";
import './App.css'

interface downloadableProps {
    userData : leaveObject[], 
    leaveType : string
}

function Downloadable({userData, leaveType} : downloadableProps){

    // const [downloadLink, setDownloadLink] = useState('')

    const name =
    userData.length > 0
        ? `${userData[0].__EMPTY_1} ${leaveType}`
        : "";

        // creates individual leave sheet
    const createLeaveSheet = useCallback(async()=>{
        const actualArray = userData;

        console.log("triggered")

        const response = await fetch("/templates/LEAVE-TEMPLATE.xlsx");
    
        const arrayBuffer = await response.arrayBuffer();
    
        const workbook = new ExcelJS.Workbook();
    
        await workbook.xlsx.load(arrayBuffer);
    
        const sheet = workbook.getWorksheet("Blank Format");
         
        // console.log(sheet)
    
        if(!sheet) return
    
        sheet.getCell("C2").value = `${actualArray[0].__EMPTY_1}`;
        sheet.getCell("C4").value = `${actualArray[0].__EMPTY}`;
        sheet.getCell("C5").value = leaveType;
    
        let i = 11
    
        actualArray.forEach((element : leaveObject) => {
            // sheet.getCell(`B${i}`).value = new Date(element.__EMPTY_2).toISOString().split("T")[0];
            sheet.getCell(`B${i}`).value = new Date(element.__EMPTY_2);
            i++
        })
    
        // console.log(actualArray)
    
        const buffer = await workbook.xlsx.writeBuffer();   
    
        const blob = new Blob([buffer], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        });
    
        const url = URL.createObjectURL(blob);

        const a = document.createElement("a");  
        a.href = url;
        a.download = `${actualArray[0].__EMPTY_1}.xlsx`;
        a.click()

        URL.revokeObjectURL(url); 
        // setDownloadLink(url)
      }, [leaveType, userData])

    
    // const handleDownload = useCallback( async()=>{
    //     console.log("clicked")   
    //     console.log(downloadLink)
    //     const actualArray =  userData;
    //     const a = document.createElement("a");
    //     a.href = downloadLink;
    //     a.download = `${actualArray[0].__EMPTY_1}.xlsx`;
    //     a.click()
    // },[downloadLink, userData])


    return(
        // <></>
        <div 
            onClick={async () => {
                await createLeaveSheet()
            }}
            className="download-card"
            style={{ backgroundColor : "gray"}}
            >
            <h4>
                {name}                
            </h4>

            
            {/* {`${userData[0].__EMPTY_1}.xlsx`} */}
        </div>
    )
}

export default Downloadable

type leaveObject = {
  __EMPTY : string;
  __EMPTY_1 : string;
  __EMPTY_2 : Date;
  __EMPTY_3 : Date;
  __EMPTY_5 : number;
  __EMPTY_6 : string
}