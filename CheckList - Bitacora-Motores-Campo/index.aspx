<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Bit&aacute;cora Diaria - Cerrej&oacute;n SGIA</title>
    
    <script crossorigin src="https://unpkg.com/react@18/umd/react.production.min.js"></script>
    <script crossorigin src="https://unpkg.com/react-dom@18/umd/react-dom.production.min.js"></script>
    <script src="https://unpkg.com/@babel/standalone/babel.min.js"></script>
    <script src="https://cdn.tailwindcss.com"></script>
    
    <script>
        tailwind.config = {
            theme: {
                extend: {
                    colors: {
                        cerrejon: {
                            gold: '#E2B53C',
                            orange: '#C77953',
                            dark: '#1a202c',
                        }
                    }
                }
            }
        }
    </script>
    
    <style>
        body, html {
            margin: 0; padding: 0; min-height: 100%;
            font-family: 'Inter', 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background-image: url('./img/img-background.png'); 
            background-color: #1a202c;
            background-size: cover; 
            background-position: center center;
            background-attachment: fixed; 
            background-repeat: no-repeat;
        }
        .no-scrollbar::-webkit-scrollbar { display: none; }
        .no-scrollbar { -ms-overflow-style: none; scrollbar-width: none; }
        
        @keyframes fadeIn { from { opacity: 0; } to { opacity: 1; } }
        .animate-fade-in { animation: fadeIn 0.2s ease-out forwards; }
    </style>
</head>
<body class="antialiased text-gray-800">
    <div id="root"></div>

    <script type="text/babel">
        const { useState, useEffect, useRef } = React;

        const SP_CONFIG = {
            siteUrl: "https://glencore.sharepoint.com/sites/co-lmn-sgia/checklist",
            listTitle: "Bitacora-Motores-Campo",
            imagesListTitle: "Bitacora-user-images"
        };

        const initialFormState = {
            Title: "", 
            Fecha: "",
            OT: "",
            Turno: "1",
            Equipo: "",
            Diagnostico: "",
            Tecnicos: "",
            BacklogRegistrado: "",
            RegistraBacklogAMTFormato: "SI",
            EquipoDisponible: "SI",
            Pendientes: ""
        };

        // Utilidad para convertir un Archivo a string Base64
        const fileToBase64 = (file) => new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.readAsDataURL(file);
            reader.onload = () => resolve(reader.result);
            reader.onerror = error => reject(error);
        });

        function App() {
            const [formData, setFormData] = useState(initialFormState);
            const [evidenceFiles, setEvidenceFiles] = useState([]); 
            const [items, setItems] = useState([]);
            const [loading, setLoading] = useState(false);
            const [error, setError] = useState(null);
            const [editingId, setEditingId] = useState(null);
            
            const [modalImages, setModalImages] = useState(null); 
            const [activeImageIndex, setActiveImageIndex] = useState(0);
            const [loadingImages, setLoadingImages] = useState(false);

            const fileInputRef = useRef(null);

            useEffect(() => { fetchItems(); }, []);

            const getRequestDigest = async () => {
                const response = await fetch(`${SP_CONFIG.siteUrl}/_api/contextinfo`, {
                    method: 'POST',
                    headers: { "Accept": "application/json;odata=verbose" }
                });
                const data = await response.json();
                return data.d.GetContextWebInformation.FormDigestValue;
            };

            const getEntityType = async (listName) => {
                const response = await fetch(`${SP_CONFIG.siteUrl}/_api/web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`, {
                    headers: { "Accept": "application/json;odata=verbose" }
                });
                const data = await response.json();
                return data.d.ListItemEntityTypeFullName;
            };

            const fetchItems = async () => {
                setLoading(true);
                try {
                    const response = await fetch(`${SP_CONFIG.siteUrl}/_api/web/lists/getbytitle('${SP_CONFIG.listTitle}')/items?$top=100&$orderby=Created desc`, {
                        headers: { "Accept": "application/json;odata=verbose" }
                    });
                    if (!response.ok) throw new Error("Error en la red al consultar la bit\u00E1cora");
                    const data = await response.json();
                    setItems(data.d.results);
                } catch (err) {
                    setError("Fallo al cargar registros principales. " + err.message);
                } finally {
                    setLoading(false);
                }
            };

            const handleFileChange = (e) => {
                const files = Array.from(e.target.files);
                setEvidenceFiles(files);
            };

            const handleChange = (e) => {
                setFormData({ ...formData, [e.target.name]: e.target.value });
            };

            // Lógica para guardar en la lista secundaria de imágenes en Base64
            const uploadBase64Images = async (recordId, digest, entityTypeImages) => {
                if (evidenceFiles.length === 0) return;
                
                // Ejecutamos secuencialmente en lugar de Promise.all para no colapsar la API con payloads gigantes
                for (let i = 0; i < evidenceFiles.length; i++) {
                    const file = evidenceFiles[i];
                    const base64String = await fileToBase64(file);
                    
                    const imagePayload = {
                        "__metadata": { "type": entityTypeImages },
                        Title: file.name,
                        ID_x002d_Registro: recordId.toString(), // Llave foránea (Verifica si el nombre interno es este)
                        Evidencia: base64String
                    };

                    const response = await fetch(`${SP_CONFIG.siteUrl}/_api/web/lists/getbytitle('${SP_CONFIG.imagesListTitle}')/items`, {
                        method: "POST",
                        headers: {
                            "Accept": "application/json;odata=verbose",
                            "Content-Type": "application/json;odata=verbose",
                            "X-RequestDigest": digest
                        },
                        body: JSON.stringify(imagePayload)
                    });

                    if (!response.ok) {
                        const err = await response.json();
                        console.error("Error guardando imagen:", err);
                        throw new Error("Fallo al guardar la imagen en Base64. Posible exceso de tama\u00F1o (Payload Too Large).");
                    }
                }
            };

            const handleSubmit = async (e) => {
                e.preventDefault();
                setLoading(true);
                setError(null);

                try {
                    const digest = await getRequestDigest();
                    const entityTypeMain = await getEntityType(SP_CONFIG.listTitle);
                    const entityTypeImages = evidenceFiles.length > 0 ? await getEntityType(SP_CONFIG.imagesListTitle) : null;
                    
                    let formattedDate = null;
                    if (formData.Fecha) {
                        formattedDate = new Date(`${formData.Fecha}T12:00:00Z`).toISOString();
                    }

                    const itemPayload = {
                        "__metadata": { "type": entityTypeMain },
                        Title: formData.OT || "Sin OT",
                        Fecha: formattedDate,
                        OT: formData.OT,
                        Turno: formData.Turno,
                        Equipo: formData.Equipo,
                        Diagnostico: formData.Diagnostico,
                        Tecnicos: formData.Tecnicos,
                        BacklogRegistrado: formData.BacklogRegistrado,
                        RegistraBacklogAMTFormato: formData.RegistraBacklogAMTFormato,
                        EquipoDisponible: formData.EquipoDisponible,
                        Pendientes: formData.Pendientes
                    };

                    const url = editingId 
                        ? `${SP_CONFIG.siteUrl}/_api/web/lists/getbytitle('${SP_CONFIG.listTitle}')/items(${editingId})`
                        : `${SP_CONFIG.siteUrl}/_api/web/lists/getbytitle('${SP_CONFIG.listTitle}')/items`;

                    const headers = {
                        "Accept": "application/json;odata=verbose",
                        "Content-Type": "application/json;odata=verbose",
                        "X-RequestDigest": digest,
                    };

                    if (editingId) {
                        headers["IF-MATCH"] = "*";
                        headers["X-HTTP-Method"] = "MERGE";
                    }

                    const response = await fetch(url, {
                        method: "POST",
                        headers: headers,
                        body: JSON.stringify(itemPayload)
                    });

                    if (!response.ok) throw new Error("Error al guardar el registro principal.");
                    
                    let targetItemId = editingId;
                    if (!editingId) {
                        const responseData = await response.json();
                        targetItemId = responseData.d.Id;
                    }

                    // Inyectar imágenes en la lista relacional
                    if (evidenceFiles.length > 0) {
                        await uploadBase64Images(targetItemId, digest, entityTypeImages);
                    }
                    
                    setFormData(initialFormState);
                    setEditingId(null);
                    setEvidenceFiles([]);
                    if(fileInputRef.current) fileInputRef.current.value = "";
                    
                    await fetchItems();
                } catch (err) {
                    setError(err.message);
                } finally {
                    setLoading(false);
                }
            };

            const handleEdit = (item) => {
                setEditingId(item.Id);
                let parsedDate = item.Fecha ? item.Fecha.split('T')[0] : "";
                setFormData({
                    Title: item.Title || "",
                    Fecha: parsedDate,
                    OT: item.OT || "", 
                    Turno: item.Turno || "1",
                    Equipo: item.Equipo || "",
                    Diagnostico: item.Diagnostico || "",
                    Tecnicos: item.Tecnicos || "",
                    BacklogRegistrado: item.BacklogRegistrado || "",
                    RegistraBacklogAMTFormato: item.RegistraBacklogAMTFormato || "SI",
                    EquipoDisponible: item.EquipoDisponible || "SI",
                    Pendientes: item.Pendientes || ""
                });
                setEvidenceFiles([]);
                if(fileInputRef.current) fileInputRef.current.value = "";
                window.scrollTo({ top: 0, behavior: 'smooth' });
            };

            // Consulta las imágenes de la base secundaria SOLO cuando se hace clic en el ojo (Lazy Load)
            const openModal = async (recordId) => {
                setLoadingImages(true);
                try {
                    const response = await fetch(`${SP_CONFIG.siteUrl}/_api/web/lists/getbytitle('${SP_CONFIG.imagesListTitle}')/items?$filter=ID_x002d_Registro eq '${recordId}'`, {
                        headers: { "Accept": "application/json;odata=verbose" }
                    });
                    if (!response.ok) throw new Error("Error obteniendo im\u00E1genes");
                    const data = await response.json();
                    
                    if (data.d.results.length === 0) {
                        alert("Este registro no tiene im\u00E1genes vinculadas o el nombre interno 'ID_x002d_Registro' es incorrecto.");
                        return;
                    }
                    
                    setModalImages(data.d.results);
                    setActiveImageIndex(0);
                } catch (err) {
                    alert("Fallo la carga de im\u00E1genes. Verifica la consola.");
                    console.error(err);
                } finally {
                    setLoadingImages(false);
                }
            };

            const closeModal = () => {
                setModalImages(null);
            };

            const prevImage = () => {
                setActiveImageIndex((prev) => (prev === 0 ? modalImages.length - 1 : prev - 1));
            };

            const nextImage = () => {
                setActiveImageIndex((prev) => (prev === modalImages.length - 1 ? 0 : prev + 1));
            };

            const glassCard = "bg-white/70 backdrop-blur-xl border border-white/40 shadow-2xl rounded-2xl";
            const inputClass = "block w-full rounded-xl bg-white/60 border-white/50 shadow-inner focus:bg-white focus:border-cerrejon-orange focus:ring-2 focus:ring-cerrejon-orange/50 transition-all duration-300 p-3 text-sm font-medium outline-none";
            const labelClass = "block text-xs font-bold text-gray-700 mb-2 uppercase tracking-widest";

            return (
                <div className="max-w-7xl mx-auto py-10 px-4 sm:px-6 lg:px-8 min-h-screen flex flex-col gap-10 relative">
                    
                    <header className={`${glassCard} p-6 flex flex-col md:flex-row items-center justify-between relative overflow-hidden`}>
                        <div className="absolute top-0 left-0 w-full h-1 bg-gradient-to-r from-cerrejon-gold via-cerrejon-orange to-red-600"></div>
                        <div className="flex items-center gap-5 z-10">
                            <div className="p-3 bg-white/80 rounded-xl shadow-sm backdrop-blur-sm">
                                <svg width="36" height="36" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                                    <path d="M12 2L2 22H22L12 2Z" fill="#C77953"/>
                                    <path d="M12 8L6 20H18L12 8Z" fill="#E2B53C"/>
                                    <path d="M12 14L9 20H15L12 14Z" fill="#1a202c"/>
                                </svg>
                            </div>
                            <div>
                                <h1 className="text-3xl font-black text-gray-900 tracking-tight">Cerrej&oacute;n <span className="text-cerrejon-orange font-light">SGIA</span></h1>
                                <p className="text-xs font-bold text-gray-600 uppercase tracking-[0.2em] mt-1">Bit&aacute;cora de Motores</p>
                            </div>
                        </div>
                    </header>
                    
                    {error && (
                        <div className="bg-red-500/90 backdrop-blur-md text-white border-l-4 border-white p-4 rounded-xl shadow-lg animate-pulse">
                            <p className="font-medium tracking-wide text-sm">{error}</p>
                        </div>
                    )}

                    {/* FORMULARIO: AHORA OCUPA TODO EL ANCHO ARRIBA */}
                    <div className={`${glassCard} p-8 w-full transition-all duration-500`}>
                        <div className="mb-6 flex items-center justify-between">
                            <h2 className="text-2xl font-extrabold text-gray-800 tracking-tight">
                                {editingId ? "Actualizar Registro Existente" : "Nuevo Registro en Bit\u00E1cora"}
                            </h2>
                            {editingId && <span className="px-4 py-2 bg-cerrejon-gold/20 text-cerrejon-orange text-sm font-bold rounded-full uppercase">Modo Edici&oacute;n</span>}
                        </div>

                        <form onSubmit={handleSubmit} className="space-y-6">
                            
                            <div className="grid grid-cols-1 md:grid-cols-4 gap-5">
                                <div>
                                    <label className={labelClass}>Fecha de Ejecuci&oacute;n</label>
                                    <input type="date" name="Fecha" value={formData.Fecha} onChange={handleChange} className={inputClass} required />
                                </div>
                                <div>
                                    <label className={labelClass}>N&uacute;mero OT</label>
                                    <input type="text" name="OT" value={formData.OT} onChange={handleChange} className={inputClass} placeholder="Ej. 104599" required />
                                </div>
                                <div>
                                    <label className={labelClass}>Turno Asignado</label>
                                    <select name="Turno" value={formData.Turno} onChange={handleChange} className={inputClass}>
                                        <option value="1">Turno 1 (Ma&ntilde;ana)</option>
                                        <option value="2">Turno 2 (Tarde)</option>
                                        <option value="3">Turno 3 (Noche)</option>
                                    </select>
                                </div>
                                <div>
                                    <label className={labelClass}>Equipo Intervenido</label>
                                    <input type="text" name="Equipo" value={formData.Equipo} onChange={handleChange} className={inputClass} placeholder="ID del Equipo" required />
                                </div>
                            </div>

                            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                                <div className="space-y-6">
                                    <div>
                                        <label className={labelClass}>Diagn&oacute;stico Inicial</label>
                                        <textarea name="Diagnostico" value={formData.Diagnostico} onChange={handleChange} rows="3" className={`${inputClass} resize-none`} placeholder="Describe brevemente la falla..."></textarea>
                                    </div>
                                    <div>
                                        <label className={labelClass}>T&eacute;cnicos a Cargo</label>
                                        <input type="text" name="Tecnicos" value={formData.Tecnicos} onChange={handleChange} className={inputClass} placeholder="Nombres separados por coma" />
                                    </div>
                                    <div className="grid grid-cols-2 gap-5 bg-white/30 p-4 rounded-xl border border-white/50">
                                        <div>
                                            <label className={labelClass}>&iquest;Registra Backlog?</label>
                                            <select name="RegistraBacklogAMTFormato" value={formData.RegistraBacklogAMTFormato} onChange={handleChange} className={inputClass}>
                                                <option value="SI">S&iacute;, Registrado</option>
                                                <option value="NO">No Requiere</option>
                                            </select>
                                        </div>
                                        <div>
                                            <label className={labelClass}>Estado del Equipo</label>
                                            <select name="EquipoDisponible" value={formData.EquipoDisponible} onChange={handleChange} className={inputClass}>
                                                <option value="SI">Disponible (Operativo)</option>
                                                <option value="NO">No Disponible (Parado)</option>
                                            </select>
                                        </div>
                                    </div>
                                </div>
                                
                                <div className="space-y-6">
                                    <div>
                                        <label className={labelClass}>Detalle del Backlog</label>
                                        <textarea name="BacklogRegistrado" value={formData.BacklogRegistrado} onChange={handleChange} rows="3" className={`${inputClass} resize-none`} placeholder="Especificaciones del backlog..."></textarea>
                                    </div>
                                    <div>
                                        <label className={labelClass}>Trabajos Pendientes</label>
                                        <textarea name="Pendientes" value={formData.Pendientes} onChange={handleChange} rows="3" className={`${inputClass} resize-none`} placeholder="Repuestos, firmas, aprobaciones..."></textarea>
                                    </div>
                                    
                                    <div className="bg-white/40 p-5 rounded-xl border border-dashed border-gray-400">
                                        <label className={labelClass}>Adjuntar Im&aacute;genes (Para convertir a Base64)</label>
                                        <input 
                                            type="file" 
                                            accept="image/*" 
                                            multiple 
                                            onChange={handleFileChange} 
                                            ref={fileInputRef}
                                            className="block w-full text-sm text-gray-700 file:mr-4 file:py-2 file:px-4 file:rounded-full file:border-0 file:text-sm file:font-bold file:bg-cerrejon-orange file:text-white hover:file:bg-orange-800 transition-colors cursor-pointer"
                                        />
                                        {evidenceFiles.length > 0 && (
                                            <p className="mt-2 text-xs font-bold text-red-600 bg-red-100 p-2 rounded">
                                                Advertencia: Estas {evidenceFiles.length} im&aacute;genes se convertir&aacute;n a texto. Si pesan mucho, la red rechazar&aacute; la petici&oacute;n.
                                            </p>
                                        )}
                                    </div>
                                </div>
                            </div>

                            <div className="flex justify-end gap-3 pt-4 border-t border-white/30">
                                {editingId && (
                                    <button type="button" onClick={() => {
                                        setEditingId(null); 
                                        setFormData(initialFormState);
                                        setEvidenceFiles([]);
                                        if(fileInputRef.current) fileInputRef.current.value = "";
                                    }} className="px-6 py-3 bg-white/50 text-gray-700 font-bold rounded-xl hover:bg-white/80 transition-colors">
                                        Cancelar Edici&oacute;n
                                    </button>
                                )}
                                <button type="submit" disabled={loading} className="px-10 py-3 bg-gradient-to-r from-cerrejon-orange to-red-600 text-white font-bold rounded-xl shadow-lg hover:shadow-xl hover:scale-[1.02] transition-all disabled:opacity-50 disabled:scale-100">
                                    {loading ? "Codificando y Guardando..." : (editingId ? "Actualizar Registro" : "Crear Registro y Subir Base64")}
                                </button>
                            </div>
                        </form>
                    </div>

                    {/* TABLA DE REGISTROS: ABAJO Y A LO ANCHO */}
                    <div className={`${glassCard} flex flex-col w-full overflow-hidden`}>
                        <div className="bg-gray-900/80 backdrop-blur-md px-6 py-5 flex justify-between items-center">
                            <h2 className="text-lg font-bold text-white uppercase tracking-widest">Base de Datos de Motores</h2>
                            <span className="bg-cerrejon-orange/20 text-cerrejon-gold text-xs px-3 py-1 rounded-md font-bold">{items.length} Registros Encontrados</span>
                        </div>
                        
                        <div className="overflow-x-auto p-4">
                            <table className="w-full text-left border-collapse min-w-[800px]">
                                <thead>
                                    <tr className="border-b border-gray-400/50 text-xs uppercase font-extrabold text-gray-800">
                                        <th className="p-4">OT / Equipo</th>
                                        <th className="p-4">Fecha & Turno</th>
                                        <th className="p-4">Diagn&oacute;stico</th>
                                        <th className="p-4">T&eacute;cnicos</th>
                                        <th className="p-4 text-center">Estado</th>
                                        <th className="p-4 text-center">Acciones</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {items.length === 0 ? (
                                        <tr><td colSpan="6" className="text-center p-8 text-gray-600 font-medium">No hay bit&aacute;coras registradas.</td></tr>
                                    ) : (
                                        items.map((item) => {
                                            const isAvailable = item.EquipoDisponible === 'SI';
                                            return (
                                                <tr key={item.Id} className="border-b border-gray-300/30 hover:bg-white/50 transition-colors">
                                                    <td className="p-4">
                                                        <div className="font-black text-cerrejon-orange text-lg">OT-{item.OT || item.Title}</div>
                                                        <div className="font-bold text-gray-800">{item.Equipo}</div>
                                                    </td>
                                                    <td className="p-4">
                                                        <div className="font-bold text-gray-700">{item.Fecha ? item.Fecha.split('T')[0] : "N/A"}</div>
                                                        <div className="text-xs text-gray-500 font-medium uppercase">Turno {item.Turno}</div>
                                                    </td>
                                                    <td className="p-4 text-sm text-gray-700 max-w-xs truncate" title={item.Diagnostico}>
                                                        {item.Diagnostico || "-"}
                                                    </td>
                                                    <td className="p-4 text-sm font-medium text-gray-600">
                                                        {item.Tecnicos || "-"}
                                                    </td>
                                                    <td className="p-4 text-center">
                                                        <span className={`px-3 py-1 rounded text-xs font-bold shadow-sm ${isAvailable ? 'bg-green-100 text-green-800' : 'bg-red-100 text-red-800'}`}>
                                                            {isAvailable ? 'Operativo' : 'Fuera de Servicio'}
                                                        </span>
                                                    </td>
                                                    <td className="p-4">
                                                        <div className="flex items-center justify-center gap-3">
                                                            <button 
                                                                onClick={() => openModal(item.Id)} 
                                                                title="Consultar im\u00E1genes Base64 en BD secundaria" 
                                                                disabled={loadingImages}
                                                                className="text-white bg-cerrejon-dark hover:bg-gray-700 p-2 rounded transition-colors flex items-center justify-center shadow-md disabled:opacity-50"
                                                            >
                                                                <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" viewBox="0 0 20 20" fill="currentColor">
                                                                    <path d="M10 12a2 2 0 100-4 2 2 0 000 4z" />
                                                                    <path fillRule="evenodd" d="M.458 10C1.732 5.943 5.522 3 10 3s8.268 2.943 9.542 7c-1.274 4.057-5.064 7-9.542 7S1.732 14.057.458 10zM14 10a4 4 0 11-8 0 4 4 0 018 0z" clipRule="evenodd" />
                                                                </svg>
                                                            </button>
                                                            <button 
                                                                onClick={() => handleEdit(item)} 
                                                                title="Editar Registro" 
                                                                className="text-cerrejon-orange bg-orange-100 hover:bg-orange-200 p-2 rounded transition-colors shadow-md"
                                                            >
                                                                <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" viewBox="0 0 20 20" fill="currentColor">
                                                                    <path d="M13.586 3.586a2 2 0 112.828 2.828l-.793.793-2.828-2.828.793-.793zM11.379 5.793L3 14.172V17h2.828l8.38-8.379-2.83-2.828z" />
                                                                </svg>
                                                            </button>
                                                        </div>
                                                    </td>
                                                </tr>
                                            );
                                        })
                                    )}
                                </tbody>
                            </table>
                        </div>
                    </div>

                    {/* LIGHTBOX DE EVIDENCIAS BASE64 */}
                    {modalImages && modalImages.length > 0 && (
                        <div className="fixed inset-0 z-[100] flex flex-col bg-black/95 backdrop-blur-md animate-fade-in">
                            <div className="flex justify-between items-center p-4 text-white border-b border-white/10">
                                <div className="font-bold tracking-widest text-sm text-cerrejon-gold uppercase">
                                    Evidencia {activeImageIndex + 1} de {modalImages.length}
                                </div>
                                <div className="flex gap-4 items-center">
                                    <span className="text-xs text-gray-400 hidden md:block">{modalImages[activeImageIndex].Title}</span>
                                    <button onClick={closeModal} className="text-white hover:text-red-500 transition-colors bg-white/10 p-2 rounded-full">
                                        <svg xmlns="http://www.w3.org/2000/svg" className="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
                                        </svg>
                                    </button>
                                </div>
                            </div>

                            <div className="flex-1 flex items-center justify-between relative p-4 overflow-hidden">
                                {modalImages.length > 1 && (
                                    <button onClick={prevImage} className="absolute left-4 z-10 p-3 bg-black/50 hover:bg-cerrejon-orange text-white rounded-full transition-all">
                                        <svg xmlns="http://www.w3.org/2000/svg" className="h-8 w-8" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M15 19l-7-7 7-7" />
                                        </svg>
                                    </button>
                                )}
                                
                                <div className="w-full h-full flex items-center justify-center">
                                    <img 
                                        src={modalImages[activeImageIndex].Evidencia} 
                                        alt={modalImages[activeImageIndex].Title} 
                                        className="max-h-full max-w-full object-contain drop-shadow-2xl" 
                                    />
                                </div>

                                {modalImages.length > 1 && (
                                    <button onClick={nextImage} className="absolute right-4 z-10 p-3 bg-black/50 hover:bg-cerrejon-orange text-white rounded-full transition-all">
                                        <svg xmlns="http://www.w3.org/2000/svg" className="h-8 w-8" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 5l7 7-7 7" />
                                        </svg>
                                    </button>
                                )}
                            </div>

                            {modalImages.length > 1 && (
                                <div className="h-28 bg-black/80 flex items-center overflow-x-auto gap-2 p-4 no-scrollbar border-t border-white/10">
                                    {modalImages.map((img, idx) => (
                                        <div 
                                            key={idx} 
                                            onClick={() => setActiveImageIndex(idx)}
                                            className={`h-full min-w-[5rem] md:min-w-[7rem] cursor-pointer rounded-lg overflow-hidden transition-all duration-300 border-2 ${idx === activeImageIndex ? 'border-cerrejon-orange scale-105 opacity-100' : 'border-transparent opacity-40 hover:opacity-80'}`}
                                        >
                                            <img src={img.Evidencia} alt="miniatura" className="w-full h-full object-cover" />
                                        </div>
                                    ))}
                                </div>
                            )}
                        </div>
                    )}

                </div>
            );
        }

        const root = ReactDOM.createRoot(document.getElementById('root'));
        root.render(<App />);
    </script>
</body>
</html>
