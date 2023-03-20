import { useState } from 'react'
import './App.scss'
import { UploadIcon, DownloadIcon } from '@primer/octicons-react'
import axios from 'axios'
import fileDownload from 'js-file-download'

function App() {
    const [archiveType, setArchiveType] = useState("")
    const [selectedFile, setSelectedFile] = useState(null)
    const [downloadUrl, setDownloadUrl] = useState("");
    const [loading, setLoading] = useState(false)
    const [download, setDownload] = useState(false)
    const [errorCode, setErrorCode] = useState("")

    const handleFileInputChange = (e) => {
        setSelectedFile(e.target.files[0])
    }
    const handleFileSubmit = (e) => {
        e.preventDefault()
        const formData = new FormData()
        formData.append('file', selectedFile)
        console.log(formData)
    }

    function extractExtention(archiveName) {
        const extention = archiveName.name.toLowerCase().split(".");
        return extention[extention.length - 1];
    }

    const handleClick = async () => {
        if (archiveType !== "") {
            setLoading(true)
            setDownload(false)
            setErrorCode("")
            let downloadFile = null; // Definir la variable antes del bloque try
            const apiURL = "http://localhost:8000/api/v1/convert-" + extractExtention(selectedFile) + "-to-" + archiveType;
            const formData = new FormData();
            formData.append('file', selectedFile);
            try {
                const response = await axios.post(apiURL, formData, {
                    headers: {
                        "Content-Type": "multipart/form-data",
                    },
                    responseType: "blob",
                });
                downloadFile =  new Blob([response.data]);
                return downloadFile
            } catch (error) {
                const statusCode = error.response?.status ?? 500; // Usamos el operador de encadenamiento opcional y de fusión nula para manejar el caso en que error.response sea undefined
                setErrorCode(statusCode)
                
            } finally {
                setLoading(false)
                setDownloadUrl(downloadFile);
                setDownload(true)
            }
        }
    }

    function checkFileExtention(file) {
        if (file.name.toLowerCase().endsWith('.pdf') || file.name.endsWith('.csv') || file.name.endsWith('.xsl') || file.name.endsWith('.doc') || file.name.endsWith('.docx')) {
            return true
        } else {
            return false
        }
    }

    function selectOptions(file) {
        let options = ["pdf", "csv", "html", "xls", "xlsx", "doc", "docx"];
        let extention = extractExtention(file);
        options.map((option, i) => {
            if (option === extention){
                options.splice(i, 1)
            }
            if ((extention === "pdf" && option === "html") /*|| (extention === "pdf" && option === "doc")*/){
                options.splice(i, 1)
            }
        })

        return options.map((option) => {
            return (
                <option key={option} value={option} id={option}>
                    {option.toUpperCase()}
                </option>
            )
        })
    }

    function getFileName(str) {
        const fileName = str.split('.')[0];
        return fileName;
    }

    function loadingFunc(boolean){
        if (boolean) {
            return (
                <span className="py-2 px-4 flex justify-center items-center  bg-blue-600 text-white transition ease-in duration-200 text-center text-base font-semibold shadow-md rounded-lg w-36 mx-auto">
                    <svg width="20" height="20" fill="currentColor" className="mr-2 animate-spin" viewBox="0 0 1792 1792" xmlns="http://www.w3.org/2000/svg">
                        <path d="M526 1394q0 53-37.5 90.5t-90.5 37.5q-52 0-90-38t-38-90q0-53 37.5-90.5t90.5-37.5 90.5 37.5 37.5 90.5zm498 206q0 53-37.5 90.5t-90.5 37.5-90.5-37.5-37.5-90.5 37.5-90.5 90.5-37.5 90.5 37.5 37.5 90.5zm-704-704q0 53-37.5 90.5t-90.5 37.5-90.5-37.5-37.5-90.5 37.5-90.5 90.5-37.5 90.5 37.5 37.5 90.5zm1202 498q0 52-38 90t-90 38q-53 0-90.5-37.5t-37.5-90.5 37.5-90.5 90.5-37.5 90.5 37.5 37.5 90.5zm-964-996q0 66-47 113t-113 47-113-47-47-113 47-113 113-47 113 47 47 113zm1170 498q0 53-37.5 90.5t-90.5 37.5-90.5-37.5-37.5-90.5 37.5-90.5 90.5-37.5 90.5 37.5 37.5 90.5zm-640-704q0 80-56 136t-136 56-136-56-56-136 56-136 136-56 136 56 56 136zm530 206q0 93-66 158.5t-158 65.5q-93 0-158.5-65.5t-65.5-158.5q0-92 65.5-158t158.5-66q92 0 158 66t66 158z">
                        </path>
                    </svg>
                    loading
                </span>
            )
        } else {
            return (<span></span>)
        }
    }

    function downloadButton(boolean) {
        if (boolean) {
            return (
                <button onClick={() => {fileDownload(downloadUrl, `${getFileName(selectedFile.name)}.zip`)}}> <DownloadIcon size={24}/> Descargar</button>
            )
        } else {
            return (<span></span>)
        }
    }

    function errorMessages(error_Code) {
        if (error_Code) {
            return (
                <h2 className='py-2'>Error {error_Code}</h2>
            )
        } else {
            return (<span></span>)
        }
    }

    return (
        <div className='App'>
            <form id='formButton' onSubmit={handleFileSubmit} className='py-2'>
                <label id='labelButton' htmlFor="fileInput" className="py-2 px-4 flex w-40 justify-center items-center bg-red-600 text-white transition ease-in duration-200 text-center text-base font-semibold shadow-md  rounded-lg mx-auto">
                    <span className='mr-2 transition'>
                        <UploadIcon size={24} id="uploadIcon"/>
                    </span>
                    Subir archivo
                </label>
                <input type="file" id='fileInput' accept='.pdf, .csv, .xsl, .doc, .docx' onChange={handleFileInputChange} className="hidden"/>
            </form>

            {selectedFile && !checkFileExtention(selectedFile) && (
                <h2 className='py-2'>El archivo de extención {extractExtention(selectedFile)} no es valido</h2>
            )}

            {selectedFile && checkFileExtention(selectedFile) && (
                <div>
                    <h2 className='py-2'>Archivo seleccionado: {selectedFile.name}</h2>
                    <select id='archiveType' className="block px-3 py-2 text-gray-700 bg-white border border-gray-300 rounded-md shadow-sm w-52 focus:outline-none focus:ring-primary-500 focus:border-primary-500 mx-auto my-2" name="animals" value={archiveType} onChange={(e) => setArchiveType(e.target.value)}>
                        <option value="">
                            Selecciona una opción
                        </option>
                        {selectOptions(selectedFile)}
                    </select>
                    <button type='submit' onClick={handleClick} className='my-2 bg-[#1a1a1a]'>Convertir</button>
                    <div>
                        {loadingFunc(loading)}
                        {errorMessages(errorCode)}
                    </div>
                </div>
            )}

            {!selectedFile && (
                <h2 className='py-2'>No hay un archivo seleccionado</h2>
            )}

            {downloadUrl &&
                downloadButton(download)
            }
        </div>
    )
}

export default App
