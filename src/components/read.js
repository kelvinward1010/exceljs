import React, { useEffect, useState } from 'react'
import * as XLSX from 'xlsx/xlsx.mjs';

export default function ReadFile() {
    const [data,setData] = useState([]);
    const [file,setFile] = useState({});
    const [error, setError] = useState('');

    const [datamain, setDataMain] = useState()

    const fileUploadButton = () => {
        document.getElementById('fileButton').click();
        document.getElementById('fileButton').onchange = () =>{      
            setFile({fileUploadState:document.getElementById('fileButton').value});
        }
    }

    const handleOpenFile = (e) => {
        const files = e.target.files;

        if (files.length) {
            const file = files[0];
            const reader = new FileReader();
            reader.onload = (event) => {
                const wb = XLSX.read(event.target.result);
                const sheets = wb.SheetNames;

                if (sheets.length) {
                    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheets[3]]);
                    setData(rows)
                }
            }
            reader.readAsArrayBuffer(file);
            
        }else{
            setError("Can't read this file")
        }
    }
    
    // useEffect(() => {
    //     setDataMain({
    //         ...datamain, 
    //         ['maSDN']: data[3]?.__EMPTY_1.substr(22), 
    //         ['maSP']: data[5]?.__EMPTY_1.substr(13)
    //     })
    // },[data])


    useEffect(() => {
        let t = {
            maSDN: data[3]?.__EMPTY_1.substr(22),
            maSP: data[5]?.__EMPTY_1.substr(13),
            nvl: data.splice(8, data.length -11)
        }
        setDataMain(t)
    },[data])
    
    console.log(datamain)
    //console.log(datamain)

    // const l = data.splice(0,8)
    // console.log(l)
    
    return (
        <>
            <div 
                className="dndnode" 
                draggable
                onClick={fileUploadButton}
            >
                Open File [CSV, XLXS]
                <input 
                    id="fileButton" 
                    type="file" 
                    onChange={handleOpenFile}
                    hidden 
                    name={'file'}
                />
            </div>
        </>
    )
}
