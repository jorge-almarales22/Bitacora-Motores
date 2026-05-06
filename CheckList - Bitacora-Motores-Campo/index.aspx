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
                        cerrejon: { gold: '#E2B53C', orange: '#C77953', dark: '#1a202c' }
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
            background-size: cover; background-position: center center;
            background-attachment: fixed; background-repeat: no-repeat;
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
            listTitle: "DB_BITACORA"
        };

        const getCurrentDate = () => new Date().toISOString().split('T')[0];

        const initialFormState = {
            Fecha: getCurrentDate(),
            OTCliente: "",
            Turno: "Dia",
            Grupo: "Guerreros",
            Equipo: "",
            Diagnostico: "",
            TrabajosRealizados: "",
            Tecnicos: [], 
            RegistraBacklogAMTFormato: "NO",
            Backlogs: [], 
            EquipoDisponible: "SI",
            HistorialPendientes: [], 
            ImagenesBase64: [] 
        };

        const fileToBase64 = (file) => new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.readAsDataURL(file);
            reader.onload = () => resolve(reader.result);
            reader.onerror = error => reject(error);
        });

        const PeoplePicker = ({ siteUrl, selectedUsers, onChange, disabled }) => {
            const [query, setQuery] = useState("");
            const [results, setResults] = useState([]);

            const searchUsers = async (searchTerm) => {
                if (searchTerm.length < 3) { setResults([]); return; }
                try {
                    const response = await fetch(`${siteUrl}/_api/contextinfo`, { method: 'POST', headers: { "Accept": "application/json;odata=verbose" }});
                    const data = await response.json();
                    const digest = data.d.GetContextWebInformation.FormDigestValue;

                    const payload = {
                        'queryParams': {
                            '__metadata': { 'type': 'SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters' },
                            'AllowEmailAddresses': true, 'AllowMultipleEntities': false, 'AllUrlZones': false,
                            'MaximumEntitySuggestions': 5, 'PrincipalSource': 15, 'PrincipalType': 1, 'QueryString': searchTerm
                        }
                    };

                    const searchRes = await fetch(`${siteUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`, {
                        method: 'POST', body: JSON.stringify(payload),
                        headers: { "Accept": "application/json;odata=verbose", "Content-Type": "application/json;odata=verbose", "X-RequestDigest": digest }
                    });
                    
                    const searchData = await searchRes.json();
                    setResults(JSON.parse(searchData.d.ClientPeoplePickerSearchUser));
                } catch (error) { console.error(error); }
            };

            useEffect(() => {
                const timeoutId = setTimeout(() => searchUsers(query), 500);
                return () => clearTimeout(timeoutId);
            }, [query]);

            const addUser = (user) => {
                if (!selectedUsers.find(u => u.correo === user.EntityData.Email)) {
                    onChange([...selectedUsers, { nombre: user.DisplayText, correo: user.EntityData.Email }]);
                }
                setQuery(""); setResults([]);
            };

            const removeUser = (correo) => { onChange(selectedUsers.filter(u => u.correo !== correo)); };

            return (
                <div className="relative">
                    <div className="flex flex-wrap gap-2 mb-2">
                        {selectedUsers.map((user, idx) => (
                            <div key={idx} className="flex items-center gap-2 bg-cerrejon-orange/20 text-cerrejon-dark px-3 py-1 rounded-full text-xs font-bold border border-cerrejon-orange/50">
                                <img src={`${siteUrl}/_layouts/15/userphoto.aspx?size=S&accountname=${user.correo}`} className="w-5 h-5 rounded-full object-cover bg-white" alt="foto" onError={(e)=>{e.target.style.display='none'}} />
                                <span>{user.nombre}</span>
                                {!disabled && <button type="button" onClick={() => removeUser(user.correo)} className="hover:text-red-600">&times;</button>}
                            </div>
                        ))}
                    </div>
                    {!disabled && (
                        <input type="text" value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Buscar nombre o correo..." className="block w-full rounded-xl bg-white/60 border-white/50 shadow-inner focus:bg-white focus:border-cerrejon-orange focus:ring-2 focus:ring-cerrejon-orange/50 transition-all p-3 text-sm outline-none" />
                    )}
                    {results.length > 0 && (
                        <ul className="absolute z-50 w-full mt-1 bg-white border border-gray-200 rounded-xl shadow-xl max-h-60 overflow-y-auto">
                            {results.map((user, idx) => (
                                <li key={idx} onClick={() => addUser(user)} className="flex items-center gap-3 p-3 hover:bg-gray-100 cursor-pointer border-b border-gray-100">
                                    <div className="flex flex-col"><span className="text-sm font-bold text-gray-800">{user.DisplayText}</span><span className="text-xs text-gray-500">{user.EntityData.Email}</span></div>
                                </li>
                            ))}
                        </ul>
                    )}
                </div>
            );
        };

        function App() {
            const [formData, setFormData] = useState(initialFormState);
            const [nuevoPendiente, setNuevoPendiente] = useState("");
            const [evidenceFiles, setEvidenceFiles] = useState([]); 
            
            const [items, setItems] = useState([]);
            const [loading, setLoading] = useState(false);
            const [error, setError] = useState(null);
            
            // Estados de bloqueo y edición
            const [editingId, setEditingId] = useState(null);
            const [originalEstado, setOriginalEstado] = useState(null);
            const [originalBacklog, setOriginalBacklog] = useState(null);
            
            const isEditing = !!editingId;
            const lockBase = isEditing; 
            const lockEstado = isEditing && originalEstado === "SI"; 
            const lockRegistraBacklog = isEditing && originalBacklog === "SI";
            const isClosed = originalEstado === "SI";
            
            const [viewModalItem, setViewModalItem] = useState(null); 
            const [modalImages, setModalImages] = useState(null); 
            const [activeImageIndex, setActiveImageIndex] = useState(0);

            // Estado para el modal dinámico de pendientes
            const [isPendingModalOpen, setIsPendingModalOpen] = useState(false);
            const [pendingModalData, setPendingModalData] = useState({ Turno: "Dia", Grupo: "Guerreros", Tecnicos: [], texto: "" });

            const fileInputRef = useRef(null);

            useEffect(() => { fetchItems(); }, []);

            const getRequestDigest = async () => {
                const response = await fetch(`${SP_CONFIG.siteUrl}/_api/contextinfo`, { method: 'POST', headers: { "Accept": "application/json;odata=verbose" }});
                const data = await response.json();
                return data.d.GetContextWebInformation.FormDigestValue;
            };

            const getEntityType = async () => {
                const response = await fetch(`${SP_CONFIG.siteUrl}/_api/web/lists/getbytitle('${SP_CONFIG.listTitle}')?$select=ListItemEntityTypeFullName`, { headers: { "Accept": "application/json;odata=verbose" }});
                const data = await response.json();
                return data.d.ListItemEntityTypeFullName;
            };

            const fetchItems = async () => {
                setLoading(true);
                try {
                    const response = await fetch(`${SP_CONFIG.siteUrl}/_api/web/lists/getbytitle('${SP_CONFIG.listTitle}')/items?$top=100&$orderby=Created desc`, { headers: { "Accept": "application/json;odata=verbose" }});
                    if (!response.ok) throw new Error("Error consultando SharePoint");
                    const data = await response.json();
                    
                    const parsedItems = data.d.results.map(item => {
                        try { return { ...item, parsedData: JSON.parse(item.Data) }; } 
                        catch (e) { return { ...item, parsedData: { Error: "Formato JSON corrupto" } }; }
                    });
                    setItems(parsedItems);
                } catch (err) { setError("Fallo al cargar registros principales."); } 
                finally { setLoading(false); }
            };

            const handleChange = (e) => { setFormData({ ...formData, [e.target.name]: e.target.value }); };

            const handleFileChange = (e) => {
                setEvidenceFiles(Array.from(e.target.files));
            };

            const addBacklog = () => {
                setFormData({
                    ...formData,
                    Backlogs: [...(formData.Backlogs || []), { id: Date.now(), diagnostico: '', hd: '', hh: '', prioridad: 'P1', repuestos: [] }]
                });
            };

            const updateBacklog = (index, field, value) => {
                const newBacklogs = [...formData.Backlogs];
                newBacklogs[index][field] = value;
                setFormData({ ...formData, Backlogs: newBacklogs });
            };

            const addRepuesto = (backlogIndex) => {
                const newBacklogs = [...formData.Backlogs];
                newBacklogs[backlogIndex].repuestos.push({ stockcode: '', cantidad: '' });
                setFormData({ ...formData, Backlogs: newBacklogs });
            };

            const updateRepuesto = (bIndex, rIndex, field, value) => {
                const newBacklogs = [...formData.Backlogs];
                newBacklogs[bIndex].repuestos[rIndex][field] = value;
                setFormData({ ...formData, Backlogs: newBacklogs });
            };

            // Guardar pendiente desde el modal
            const handleSavePendingModal = () => {
                if(!pendingModalData.texto.trim()) return;
                const newHistory = [...(formData.HistorialPendientes || [])];
                newHistory.push({
                    texto: pendingModalData.texto,
                    fechaHora: new Date().toLocaleString(),
                    turno: pendingModalData.Turno,
                    grupo: pendingModalData.Grupo,
                    tecnicos: pendingModalData.Tecnicos
                });
                setFormData({ ...formData, HistorialPendientes: newHistory });
                setIsPendingModalOpen(false);
                setPendingModalData({ Turno: "Dia", Grupo: "Guerreros", Tecnicos: [], texto: "" });
            };

            const handleSubmit = async (e) => {
                e.preventDefault();
                setLoading(true); setError(null);

                try {
                    const digest = await getRequestDigest();
                    const entityType = await getEntityType();
                    
                    let finalDataObj = { ...formData };

                    // Si NO está editando y deja un pendiente rápido en el formulario principal
                    if (!isEditing && finalDataObj.EquipoDisponible === "NO" && nuevoPendiente.trim() !== "") {
                        finalDataObj.HistorialPendientes = [...(finalDataObj.HistorialPendientes || [])];
                        finalDataObj.HistorialPendientes.push({ 
                            texto: nuevoPendiente, 
                            fechaHora: new Date().toLocaleString(),
                            turno: finalDataObj.Turno,
                            grupo: finalDataObj.Grupo,
                            tecnicos: finalDataObj.Tecnicos
                        });
                    }

                    // Concatenar pendientes en trabajos realizados si el equipo pasa a DISPONIBLE
                    if (finalDataObj.EquipoDisponible === "SI" && finalDataObj.HistorialPendientes && finalDataObj.HistorialPendientes.length > 0) {
                        const pendientesText = finalDataObj.HistorialPendientes.map(h => {
                            const tecs = h.tecnicos && h.tecnicos.length > 0 ? h.tecnicos.map(t=>t.nombre.split(' ')[0]).join(', ') : 'Sin t\u00E9cnicos';
                            return `[${h.fechaHora}] Turno: ${h.turno || 'N/A'} | Grupo: ${h.grupo || 'N/A'} | T\u00E9cs: ${tecs}\nDetalle: ${h.texto}`;
                        }).join('\n\n');
                        
                        finalDataObj.TrabajosRealizados = (finalDataObj.TrabajosRealizados ? finalDataObj.TrabajosRealizados + '\n\n' : '') + "--- TAREAS COMPLETADAS DESDE PENDIENTES ---\n" + pendientesText;
                        finalDataObj.HistorialPendientes = [];
                    }

                    let base64Images = [...(finalDataObj.ImagenesBase64 || [])];
                    for (let file of evidenceFiles) {
                        const b64 = await fileToBase64(file);
                        base64Images.push({ name: file.name, data: b64 });
                    }
                    finalDataObj.ImagenesBase64 = base64Images;

                    const itemPayload = { "__metadata": { "type": entityType }, Title: formData.OTCliente || "Sin OT", Data: JSON.stringify(finalDataObj) };
                    const url = editingId ? `${SP_CONFIG.siteUrl}/_api/web/lists/getbytitle('${SP_CONFIG.listTitle}')/items(${editingId})` : `${SP_CONFIG.siteUrl}/_api/web/lists/getbytitle('${SP_CONFIG.listTitle}')/items`;
                    const headers = { "Accept": "application/json;odata=verbose", "Content-Type": "application/json;odata=verbose", "X-RequestDigest": digest };
                    if (editingId) { headers["IF-MATCH"] = "*"; headers["X-HTTP-Method"] = "MERGE"; }

                    const response = await fetch(url, { method: "POST", headers: headers, body: JSON.stringify(itemPayload) });
                    if (!response.ok) throw new Error("Fallo al guardar. Posible exceso de tama\u00F1o por im\u00E1genes.");
                    
                    setFormData(initialFormState);
                    setNuevoPendiente(""); 
                    setEditingId(null); 
                    setOriginalEstado(null);
                    setOriginalBacklog(null);
                    setEvidenceFiles([]);
                    if(fileInputRef.current) fileInputRef.current.value = "";
                    await fetchItems();
                } catch (err) { setError(err.message); } 
                finally { setLoading(false); }
            };

            const handleEdit = (item) => {
                const pd = item.parsedData;
                if (!pd.Backlogs) pd.Backlogs = [];
                
                setOriginalEstado(pd.EquipoDisponible);
                setOriginalBacklog(pd.RegistraBacklogAMTFormato);
                setEditingId(item.Id);
                setFormData(pd);
                setNuevoPendiente(""); setEvidenceFiles([]);
                if(fileInputRef.current) fileInputRef.current.value = "";
                window.scrollTo({ top: 0, behavior: 'smooth' });
            };

            const glassCard = "bg-white/70 backdrop-blur-xl border border-white/40 shadow-2xl rounded-2xl";
            const inputClass = "block w-full rounded-xl bg-white/60 border-white/50 shadow-inner focus:bg-white focus:border-cerrejon-orange focus:ring-2 focus:ring-cerrejon-orange/50 transition-all p-3 text-sm outline-none disabled:opacity-60 disabled:cursor-not-allowed disabled:bg-gray-200/50";
            const labelClass = "block text-xs font-bold text-gray-700 mb-2 uppercase tracking-widest";

            return (
                <div className="max-w-[1400px] mx-auto py-10 px-4 sm:px-6 lg:px-8 min-h-screen flex flex-col gap-10 relative">
                    
                    <header className={`${glassCard} p-6 flex items-center justify-between relative overflow-hidden`}>
                        <div className="absolute top-0 left-0 w-full h-1 bg-gradient-to-r from-cerrejon-gold via-cerrejon-orange to-red-600"></div>
                        <div className="flex items-center gap-5 z-10">
                            <div className="p-3 bg-white/80 rounded-xl shadow-sm backdrop-blur-sm">
                                <svg width="36" height="36" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                                    <path d="M12 2L2 22H22L12 2Z" fill="#C77953"/><path d="M12 8L6 20H18L12 8Z" fill="#E2B53C"/><path d="M12 14L9 20H15L12 14Z" fill="#1a202c"/>
                                </svg>
                            </div>
                            <div>
                                <h1 className="text-3xl font-black text-gray-900 tracking-tight">Cerrej&oacute;n <span className="text-cerrejon-orange font-light">SGIA</span></h1>
                                <p className="text-xs font-bold text-gray-600 uppercase tracking-[0.2em] mt-1">Bit&aacute;cora de Motores</p>
                            </div>
                        </div>
                    </header>
                    
                    {error && (<div className="bg-red-500/90 backdrop-blur-md text-white border-l-4 border-white p-4 rounded-xl shadow-lg"><p className="font-medium text-sm">{error}</p></div>)}

                    {/* FORMULARIO DE EDICIÓN/CREACIÓN */}
                    <div className={`${glassCard} p-8 w-full transition-all duration-500`}>
                        <div className="mb-6 flex flex-col md:flex-row items-start md:items-center justify-between gap-4">
                            <h2 className="text-2xl font-extrabold text-gray-800 tracking-tight">
                                {isClosed && isEditing ? "Registro Cerrado (Solo edici\u00F3n de Backlogs)" : (isEditing ? "Actualizar Registro (Campos Base Bloqueados)" : "Nuevo Registro de Bit\u00E1cora JSON")}
                            </h2>
                            {isClosed && isEditing && <span className="px-4 py-2 bg-green-600 text-white text-xs font-bold rounded-full uppercase shadow-md">Equipo Disponible</span>}
                            {!isClosed && isEditing && <span className="px-4 py-2 bg-red-600 text-white text-xs font-bold rounded-full uppercase shadow-md">Equipo No Disponible</span>}
                        </div>

                        <form onSubmit={handleSubmit} className="space-y-6">
                            
                            <div className="grid grid-cols-1 md:grid-cols-5 gap-5">
                                <div><label className={labelClass}>Fecha de Ejecuci&oacute;n</label><input type="date" name="Fecha" value={formData.Fecha} onChange={handleChange} disabled={lockBase} className={inputClass} required /></div>
                                <div><label className={labelClass}>OT Cliente</label><input type="text" name="OTCliente" value={formData.OTCliente} onChange={handleChange} disabled={lockBase} className={inputClass} placeholder="Ej. 104599" required /></div>
                                <div>
                                    <label className={labelClass}>Turno</label>
                                    <select name="Turno" value={formData.Turno} onChange={handleChange} disabled={lockBase} className={inputClass}>
                                        <option value="Dia">D&iacute;a</option><option value="Noche">Noche</option>
                                    </select>
                                </div>
                                <div>
                                    <label className={labelClass}>Grupo</label>
                                    <select name="Grupo" value={formData.Grupo} onChange={handleChange} disabled={lockBase} className={inputClass}>
                                        <option value="Guerreros">Guerreros</option><option value="Jaguares">Jaguares</option><option value="Cardenales">Cardenales</option><option value="Valientes">Valientes</option>
                                    </select>
                                </div>
                                <div><label className={labelClass}>Equipo Intervenido</label><input type="text" name="Equipo" value={formData.Equipo} onChange={handleChange} disabled={lockBase} className={inputClass} required /></div>
                            </div>

                            <div><label className={labelClass}>Diagn&oacute;stico Inicial</label><textarea name="Diagnostico" value={formData.Diagnostico} onChange={handleChange} disabled={lockBase} rows="2" className={`${inputClass} resize-none`} placeholder="Describe brevemente la falla..."></textarea></div>

                            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                                <div className="space-y-6">
                                    <div><label className={labelClass}>T&eacute;cnicos a Cargo</label><PeoplePicker siteUrl={SP_CONFIG.siteUrl} selectedUsers={formData.Tecnicos} onChange={(users) => setFormData({...formData, Tecnicos: users})} disabled={lockBase} /></div>
                                    <div><label className={labelClass}>Trabajos Realizados</label><textarea name="TrabajosRealizados" value={formData.TrabajosRealizados} onChange={handleChange} disabled={lockBase} rows="6" className={`${inputClass} resize-none`} placeholder="Describe los trabajos ejecutados..."></textarea></div>
                                </div>
                                
                                <div className="space-y-6">
                                    <div className="grid grid-cols-2 gap-5 bg-white/30 p-4 rounded-xl border border-white/50">
                                        <div>
                                            <label className={labelClass}>Estado del Equipo</label>
                                            <select name="EquipoDisponible" value={formData.EquipoDisponible} onChange={handleChange} disabled={lockEstado} className={`${inputClass} !border-cerrejon-orange !ring-2 !ring-cerrejon-orange/50`}>
                                                <option value="SI">Disponible (Operativo)</option><option value="NO">No Disponible (Parado)</option>
                                            </select>
                                        </div>
                                        <div>
                                            <label className={labelClass}>&iquest;Registra Backlog?</label>
                                            <select name="RegistraBacklogAMTFormato" value={formData.RegistraBacklogAMTFormato} onChange={handleChange} disabled={lockRegistraBacklog} className={inputClass}>
                                                <option value="SI">S&iacute;, Registrado</option><option value="NO">No Requiere</option>
                                            </select>
                                        </div>
                                    </div>
                                    
                                    {formData.EquipoDisponible === "NO" && (
                                        <div className="bg-red-50/80 p-5 rounded-xl border border-red-200">
                                            {formData.HistorialPendientes && formData.HistorialPendientes.length > 0 && (
                                                <div className="mb-4 space-y-3">
                                                    <label className={labelClass}>Historial de Pendientes Inmutables</label>
                                                    {formData.HistorialPendientes.map((hist, i) => (
                                                        <div key={i} className="bg-white p-3 rounded shadow-sm border border-red-100 text-sm">
                                                            <div className="text-[10px] text-gray-500 font-bold mb-1 uppercase">
                                                                {hist.fechaHora} &bull; T: {hist.turno} &bull; G: {hist.grupo}
                                                            </div>
                                                            <div className="text-gray-800 mb-1">{hist.texto}</div>
                                                            <div className="text-[10px] text-gray-400 italic font-bold">
                                                                Por: {hist.tecnicos && hist.tecnicos.length > 0 ? hist.tecnicos.map(t=>t.nombre.split(' ')[0]).join(', ') : 'ND'}
                                                            </div>
                                                        </div>
                                                    ))}
                                                </div>
                                            )}
                                            
                                            {/* Si es NUEVO registro, muestra el text area normal. Si está EDITANDO, muestra el botón para el Modal */}
                                            {!isEditing ? (
                                                <div>
                                                    <label className={labelClass}>Agregar Pendiente de Cierre de Turno</label>
                                                    <textarea value={nuevoPendiente} onChange={(e) => setNuevoPendiente(e.target.value)} rows="3" className={`${inputClass} resize-none bg-white`} placeholder="Se guardar&aacute; con fecha y hora inmutable asumiendo el turno actual."></textarea>
                                                </div>
                                            ) : (
                                                !lockEstado && (
                                                    <button type="button" onClick={() => setIsPendingModalOpen(true)} className="w-full py-3 bg-red-600 hover:bg-red-700 text-white font-bold rounded-xl shadow-md transition-colors uppercase text-sm tracking-wide">
                                                        + Agregar Nuevo Trabajo Pendiente
                                                    </button>
                                                )
                                            )}
                                        </div>
                                    )}
                                    
                                    {!isEditing && (
                                        <div className="bg-white/40 p-5 rounded-xl border border-dashed border-gray-400">
                                            <label className={labelClass}>Evidencias (Codificadas a Base64)</label>
                                            <input type="file" accept="image/*" multiple onChange={handleFileChange} ref={fileInputRef} className="block w-full text-sm text-gray-700 file:mr-4 file:py-2 file:px-4 file:rounded-full file:border-0 file:font-bold file:bg-cerrejon-orange file:text-white cursor-pointer" />
                                        </div>
                                    )}
                                </div>
                            </div>

                            {/* SECCIÓN DINÁMICA DE BACKLOGS */}
                            {formData.RegistraBacklogAMTFormato === "SI" && (
                                <div className="mt-8 border-t-2 border-cerrejon-orange pt-6">
                                    <div className="flex justify-between items-center mb-4">
                                        <h3 className="text-lg font-black text-gray-800 uppercase tracking-wide">Gesti&oacute;n de Backlogs</h3>
                                        <button type="button" onClick={addBacklog} className="bg-cerrejon-dark text-white px-4 py-2 rounded font-bold text-xs hover:bg-gray-700 transition-colors">
                                            + AGREGAR BACKLOG
                                        </button>
                                    </div>
                                    
                                    {formData.Backlogs && formData.Backlogs.length === 0 && <p className="text-sm text-gray-500 italic">No hay backlogs agregados a&uacute;n.</p>}
                                    
                                    <div className="space-y-6">
                                        {formData.Backlogs && formData.Backlogs.map((backlog, bIndex) => (
                                            <div key={backlog.id} className="bg-white p-5 rounded-xl border border-gray-300 shadow-sm relative">
                                                <button type="button" onClick={() => setFormData({...formData, Backlogs: formData.Backlogs.filter((_, i) => i !== bIndex)})} className="absolute top-2 right-2 text-red-500 font-bold hover:text-red-700">&times; Eliminar</button>
                                                <div className="grid grid-cols-1 md:grid-cols-12 gap-4 mb-4">
                                                    <div className="md:col-span-6"><label className="block text-[10px] font-bold text-gray-500 uppercase">Diagn&oacute;stico</label><input type="text" value={backlog.diagnostico} onChange={e => updateBacklog(bIndex, 'diagnostico', e.target.value)} className="w-full border-b border-gray-300 outline-none focus:border-cerrejon-orange text-sm p-1" /></div>
                                                    <div className="md:col-span-2"><label className="block text-[10px] font-bold text-gray-500 uppercase">HD (Horas Down)</label><input type="number" value={backlog.hd} onChange={e => updateBacklog(bIndex, 'hd', e.target.value)} className="w-full border-b border-gray-300 outline-none focus:border-cerrejon-orange text-sm p-1" /></div>
                                                    <div className="md:col-span-2"><label className="block text-[10px] font-bold text-gray-500 uppercase">HH (Horas Hombre)</label><input type="number" value={backlog.hh} onChange={e => updateBacklog(bIndex, 'hh', e.target.value)} className="w-full border-b border-gray-300 outline-none focus:border-cerrejon-orange text-sm p-1" /></div>
                                                    <div className="md:col-span-2"><label className="block text-[10px] font-bold text-gray-500 uppercase">Prioridad</label>
                                                        <select value={backlog.prioridad} onChange={e => updateBacklog(bIndex, 'prioridad', e.target.value)} className="w-full border-b border-gray-300 outline-none focus:border-cerrejon-orange text-sm p-1 bg-transparent">
                                                            <option value="P1">P1</option><option value="P2">P2</option><option value="P3">P3</option><option value="P4">P4</option>
                                                        </select>
                                                    </div>
                                                </div>
                                                <div className="bg-gray-50 p-3 rounded border border-gray-200">
                                                    <div className="flex justify-between items-center mb-2">
                                                        <span className="text-xs font-bold text-gray-700">Lista de Repuestos</span>
                                                        <button type="button" onClick={() => addRepuesto(bIndex)} className="text-[10px] bg-gray-200 px-2 py-1 rounded font-bold hover:bg-gray-300">+ Agregar Repuesto</button>
                                                    </div>
                                                    {backlog.repuestos.map((rep, rIndex) => (
                                                        <div key={rIndex} className="flex gap-3 mb-2 items-center">
                                                            <input type="text" placeholder="Stockcode" value={rep.stockcode} onChange={e => updateRepuesto(bIndex, rIndex, 'stockcode', e.target.value)} className="w-1/2 text-xs border p-1 rounded" />
                                                            <input type="number" placeholder="Cantidad" value={rep.cantidad} onChange={e => updateRepuesto(bIndex, rIndex, 'cantidad', e.target.value)} className="w-1/3 text-xs border p-1 rounded" />
                                                            <button type="button" onClick={() => { const nb=[...formData.Backlogs]; nb[bIndex].repuestos.splice(rIndex,1); setFormData({...formData, Backlogs: nb}); }} className="text-red-500 text-lg font-bold">&times;</button>
                                                        </div>
                                                    ))}
                                                </div>
                                            </div>
                                        ))}
                                    </div>
                                </div>
                            )}

                            <div className="flex justify-end gap-3 pt-4 border-t border-white/30">
                                {isEditing && (
                                    <button type="button" onClick={() => { setEditingId(null); setOriginalEstado(null); setOriginalBacklog(null); setFormData(initialFormState); setNuevoPendiente(""); }} className="px-6 py-3 bg-white/50 text-gray-700 font-bold rounded-xl hover:bg-white/80 transition-colors">Cancelar</button>
                                )}
                                <button type="submit" disabled={loading} className="px-10 py-3 bg-gradient-to-r from-cerrejon-orange to-red-600 text-white font-bold rounded-xl shadow-lg hover:scale-[1.02] transition-all disabled:opacity-50">
                                    {loading ? "Procesando JSON..." : (isEditing ? "Actualizar Registro" : "Guardar Registro")}
                                </button>
                            </div>
                        </form>
                    </div>

                    {/* TABLA INFERIOR RESUMIDA */}
                    <div className={`${glassCard} flex flex-col w-full overflow-hidden`}>
                        <div className="bg-gray-900/80 backdrop-blur-md px-6 py-5 flex justify-between items-center">
                            <h2 className="text-lg font-bold text-white uppercase tracking-widest">Base de Datos (Vista Resumida)</h2>
                        </div>
                        
                        <div className="overflow-x-auto p-4">
                            <table className="w-full text-left border-collapse">
                                <thead>
                                    <tr className="border-b border-gray-400/50 text-xs uppercase font-extrabold text-gray-800">
                                        <th className="p-4">Fecha de Ejecuci&oacute;n</th>
                                        <th className="p-4">OT Cliente</th>
                                        <th className="p-4">Equipo Intervenido</th>
                                        <th className="p-4 w-1/3">Diagn&oacute;stico Inicial</th>
                                        <th className="p-4 text-center">Estado del Equipo</th>
                                        <th className="p-4 text-center">Acciones</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {items.length === 0 ? (
                                        <tr><td colSpan="6" className="text-center p-8 font-medium">No hay registros.</td></tr>
                                    ) : (
                                        items.map((item) => {
                                            const d = item.parsedData;
                                            if (d.Error) return (<tr key={item.Id}><td colSpan="6" className="p-4 text-red-500">ID {item.Id}: {d.Error}</td></tr>);
                                            
                                            const isAvailable = d.EquipoDisponible === 'SI';

                                            return (
                                                <tr key={item.Id} className="border-b border-gray-300/30 hover:bg-white/50 transition-colors align-top">
                                                    <td className="p-4 font-medium text-gray-800">{d.Fecha}</td>
                                                    <td className="p-4 font-black text-cerrejon-orange">OT-{d.OTCliente}</td>
                                                    <td className="p-4 font-bold text-gray-800">{d.Equipo}</td>
                                                    <td className="p-4 text-sm text-gray-700 truncate max-w-xs" title={d.Diagnostico}>{d.Diagnostico}</td>
                                                    <td className="p-4 text-center">
                                                        <span className={`px-2 py-1 rounded text-xs font-bold ${isAvailable ? 'bg-green-100 text-green-800' : 'bg-red-100 text-red-800'}`}>{isAvailable ? 'Operativo' : 'Parado'}</span>
                                                    </td>
                                                    <td className="p-4">
                                                        <div className="flex items-center justify-center gap-2">
                                                            <button onClick={() => setViewModalItem(item)} title="Ver Detalle Completo" className="text-white bg-cerrejon-dark hover:bg-gray-700 p-2 rounded shadow transition-colors">
                                                                <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" viewBox="0 0 20 20" fill="currentColor"><path d="M10 12a2 2 0 100-4 2 2 0 000 4z" /><path fillRule="evenodd" d="M.458 10C1.732 5.943 5.522 3 10 3s8.268 2.943 9.542 7c-1.274 4.057-5.064 7-9.542 7S1.732 14.057.458 10zM14 10a4 4 0 11-8 0 4 4 0 018 0z" clipRule="evenodd" /></svg>
                                                            </button>
                                                            <button onClick={() => handleEdit(item)} title="Editar" className="text-cerrejon-orange bg-orange-100 hover:bg-orange-200 p-2 rounded shadow transition-colors">
                                                                <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" viewBox="0 0 20 20" fill="currentColor"><path d="M13.586 3.586a2 2 0 112.828 2.828l-.793.793-2.828-2.828.793-.793zM11.379 5.793L3 14.172V17h2.828l8.38-8.379-2.83-2.828z" /></svg>
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

                    {/* MODAL: REGISTRAR NUEVO PENDIENTE (AL EDITAR) */}
                    {isPendingModalOpen && (
                        <div className="fixed inset-0 z-[110] flex items-center justify-center bg-black/80 backdrop-blur-sm p-4 animate-fade-in">
                            <div className="bg-white rounded-2xl w-full max-w-xl shadow-2xl overflow-hidden">
                                <div className="bg-red-600 text-white p-4 flex justify-between items-center">
                                    <h3 className="font-bold tracking-wider uppercase text-sm">Registrar Pendiente / Cambio de Turno</h3>
                                    <button onClick={() => setIsPendingModalOpen(false)} className="hover:text-gray-200">&times;</button>
                                </div>
                                <div className="p-6 space-y-5">
                                    <div className="grid grid-cols-2 gap-4">
                                        <div>
                                            <label className="block text-[10px] font-bold text-gray-500 uppercase mb-1">Turno Entrante</label>
                                            <select value={pendingModalData.Turno} onChange={(e) => setPendingModalData({...pendingModalData, Turno: e.target.value})} className="w-full border border-gray-300 rounded p-2 text-sm outline-none focus:border-red-500">
                                                <option value="Dia">D&iacute;a</option><option value="Noche">Noche</option>
                                            </select>
                                        </div>
                                        <div>
                                            <label className="block text-[10px] font-bold text-gray-500 uppercase mb-1">Grupo Entrante</label>
                                            <select value={pendingModalData.Grupo} onChange={(e) => setPendingModalData({...pendingModalData, Grupo: e.target.value})} className="w-full border border-gray-300 rounded p-2 text-sm outline-none focus:border-red-500">
                                                <option value="Guerreros">Guerreros</option><option value="Jaguares">Jaguares</option><option value="Cardenales">Cardenales</option><option value="Valientes">Valientes</option>
                                            </select>
                                        </div>
                                    </div>
                                    <div>
                                        <label className="block text-[10px] font-bold text-gray-500 uppercase mb-1">T&eacute;cnicos que reciben el pendiente</label>
                                        <PeoplePicker siteUrl={SP_CONFIG.siteUrl} selectedUsers={pendingModalData.Tecnicos} onChange={(users) => setPendingModalData({...pendingModalData, Tecnicos: users})} />
                                    </div>
                                    <div>
                                        <label className="block text-[10px] font-bold text-gray-500 uppercase mb-1">Detalle del trabajo pendiente</label>
                                        <textarea value={pendingModalData.texto} onChange={(e) => setPendingModalData({...pendingModalData, texto: e.target.value})} rows="4" className="w-full border border-gray-300 rounded p-2 text-sm outline-none focus:border-red-500 resize-none" placeholder="Especificar qu&eacute; hace falta..."></textarea>
                                    </div>
                                    <div className="flex justify-end gap-3 pt-4 border-t border-gray-100">
                                        <button type="button" onClick={() => setIsPendingModalOpen(false)} className="px-4 py-2 bg-gray-200 text-gray-700 font-bold rounded hover:bg-gray-300">Cancelar</button>
                                        <button type="button" onClick={handleSavePendingModal} className="px-6 py-2 bg-red-600 text-white font-bold rounded shadow hover:bg-red-700 disabled:opacity-50" disabled={!pendingModalData.texto.trim()}>Agregar al Historial</button>
                                    </div>
                                </div>
                            </div>
                        </div>
                    )}

                    {/* MODAL: VER DETALLE COMPLETO */}
                    {viewModalItem && (() => {
                        const d = viewModalItem.parsedData;
                        const imgs = d.ImagenesBase64 || [];
                        const isAvailable = d.EquipoDisponible === 'SI';
                        
                        return (
                            <div className="fixed inset-0 z-[90] flex items-center justify-center bg-black/80 backdrop-blur-sm animate-fade-in p-4 overflow-y-auto">
                                <div className="bg-white rounded-2xl w-full max-w-4xl shadow-2xl flex flex-col my-auto relative border-t-8 border-cerrejon-orange">
                                    
                                    <div className="flex justify-between items-center p-6 border-b border-gray-200 bg-gray-50 rounded-t-2xl">
                                        <div>
                                            <h2 className="text-2xl font-black text-gray-800">Detalle de Bit&aacute;cora</h2>
                                            <p className="text-sm font-bold text-cerrejon-orange mt-1">OT-{d.OTCliente} | Equipo: {d.Equipo}</p>
                                        </div>
                                        <button onClick={() => setViewModalItem(null)} className="text-gray-400 hover:text-red-500 transition-colors bg-white p-2 rounded-full shadow-sm border border-gray-200">
                                            <svg className="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" /></svg>
                                        </button>
                                    </div>

                                    <div className="p-6 overflow-y-auto max-h-[75vh] space-y-8">
                                        <div className="grid grid-cols-2 md:grid-cols-4 gap-4 bg-gray-100 p-4 rounded-xl border border-gray-200">
                                            <div><span className="block text-[10px] font-bold text-gray-500 uppercase">Fecha Inicio</span><span className="font-bold text-gray-800">{d.Fecha}</span></div>
                                            <div><span className="block text-[10px] font-bold text-gray-500 uppercase">Turno Inicio</span><span className="font-bold text-gray-800">{d.Turno}</span></div>
                                            <div><span className="block text-[10px] font-bold text-gray-500 uppercase">Grupo Inicio</span><span className="font-bold text-gray-800">{d.Grupo}</span></div>
                                            <div>
                                                <span className="block text-[10px] font-bold text-gray-500 uppercase">Estado General</span>
                                                <span className={`inline-block px-2 py-1 mt-1 rounded text-xs font-bold ${isAvailable ? 'bg-green-100 text-green-800' : 'bg-red-100 text-red-800'}`}>{isAvailable ? 'Operativo' : 'Parado'}</span>
                                            </div>
                                        </div>

                                        <div className="space-y-4">
                                            <div>
                                                <h4 className="text-xs font-bold text-cerrejon-dark uppercase border-b border-gray-200 pb-1 mb-2">Diagn&oacute;stico Inicial</h4>
                                                <p className="text-sm text-gray-700 bg-gray-50 p-3 rounded border border-gray-100 whitespace-pre-wrap">{d.Diagnostico || "Sin registro"}</p>
                                            </div>
                                            <div>
                                                <h4 className="text-xs font-bold text-cerrejon-dark uppercase border-b border-gray-200 pb-1 mb-2">Trabajos Realizados (Acumulado)</h4>
                                                <p className="text-sm text-gray-700 bg-gray-50 p-3 rounded border border-gray-100 whitespace-pre-wrap">{d.TrabajosRealizados || "Sin registro"}</p>
                                            </div>
                                            
                                            {!isAvailable && d.HistorialPendientes && d.HistorialPendientes.length > 0 && (
                                                <div>
                                                    <h4 className="text-xs font-bold text-red-600 uppercase border-b border-red-200 pb-1 mb-2">Historial de Pendientes y Entregas de Turno</h4>
                                                    <div className="bg-red-50 p-3 rounded border border-red-100 space-y-2">
                                                        {d.HistorialPendientes.map((h,i) => {
                                                            const tecs = h.tecnicos && h.tecnicos.length > 0 ? h.tecnicos.map(t=>t.nombre.split(' ')[0]).join(', ') : 'ND';
                                                            return (
                                                                <div key={i} className="text-sm border-b border-red-100 pb-3 last:border-0">
                                                                    <div className="flex flex-wrap gap-2 mb-1">
                                                                        <span className="font-bold text-[10px] bg-red-200 text-red-800 px-2 py-0.5 rounded">{h.fechaHora}</span>
                                                                        <span className="font-bold text-[10px] bg-white border border-red-200 text-gray-600 px-2 py-0.5 rounded">T: {h.turno || 'ND'}</span>
                                                                        <span className="font-bold text-[10px] bg-white border border-red-200 text-gray-600 px-2 py-0.5 rounded">G: {h.grupo || 'ND'}</span>
                                                                    </div>
                                                                    <div className="text-gray-800 mb-1">{h.texto}</div>
                                                                    <div className="text-[10px] text-gray-500 italic font-bold">Asignado a: {tecs}</div>
                                                                </div>
                                                            );
                                                        })}
                                                    </div>
                                                </div>
                                            )}

                                            {d.Backlogs && d.Backlogs.length > 0 && (
                                                <div>
                                                    <h4 className="text-xs font-bold text-cerrejon-dark uppercase border-b border-gray-200 pb-1 mb-2">Backlogs Registrados ({d.Backlogs.length})</h4>
                                                    <div className="grid grid-cols-1 gap-3">
                                                        {d.Backlogs.map((b,i) => (
                                                            <div key={i} className="bg-orange-50 p-3 rounded border border-orange-100 text-sm">
                                                                <div className="font-bold text-cerrejon-orange mb-1">{b.diagnostico}</div>
                                                                <div className="text-xs text-gray-600 mb-2"><b>Prio:</b> {b.prioridad} &bull; <b>HD:</b> {b.hd} &bull; <b>HH:</b> {b.hh}</div>
                                                                {b.repuestos && b.repuestos.length > 0 && (
                                                                    <div className="bg-white p-2 rounded text-xs border border-orange-50">
                                                                        <b>Repuestos:</b> {b.repuestos.map(r=>`${r.stockcode} (${r.cantidad})`).join(', ')}
                                                                    </div>
                                                                )}
                                                            </div>
                                                        ))}
                                                    </div>
                                                </div>
                                            )}
                                        </div>

                                        <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 bg-gray-100 p-4 rounded-xl border border-gray-200">
                                            <div>
                                                <h4 className="text-[10px] font-bold text-gray-500 uppercase mb-1">T&eacute;cnicos Iniciales</h4>
                                                <div className="flex flex-wrap gap-2">
                                                    {(d.Tecnicos || []).length > 0 ? d.Tecnicos.map((t, idx) => (
                                                        <span key={idx} className="bg-white border border-gray-300 text-xs font-bold px-2 py-1 rounded-full shadow-sm">{t.nombre}</span>
                                                    )) : <span className="text-xs italic text-gray-500">Ninguno</span>}
                                                </div>
                                            </div>
                                            <div>
                                                {imgs.length > 0 ? (
                                                    <button onClick={() => {setModalImages(imgs); setActiveImageIndex(0);}} className="bg-cerrejon-dark text-white font-bold text-sm px-4 py-2 rounded-lg shadow hover:bg-gray-700 transition-colors flex items-center gap-2">
                                                        <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" viewBox="0 0 20 20" fill="currentColor"><path d="M4 3a2 2 0 00-2 2v10a2 2 0 002 2h12a2 2 0 002-2V5a2 2 0 00-2-2H4zm12 12H4l4-8 3 6 2-4 3 6z" /></svg>
                                                        Ver {imgs.length} Evidencia(s)
                                                    </button>
                                                ) : (
                                                    <span className="text-xs font-bold text-gray-400 bg-gray-200 px-3 py-2 rounded-lg">Sin Evidencias Fotogr&aacute;ficas</span>
                                                )}
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        );
                    })()}

                    {/* MODAL LIGHTBOX DE IMÁGENES JSON */}
                    {modalImages && modalImages.length > 0 && (
                        <div className="fixed inset-0 z-[100] flex flex-col bg-black/95 backdrop-blur-md animate-fade-in">
                            <div className="flex justify-between items-center p-4 text-white border-b border-white/10">
                                <div className="font-bold tracking-widest text-sm text-cerrejon-gold uppercase">Evidencia {activeImageIndex + 1} de {modalImages.length}</div>
                                <button onClick={() => setModalImages(null)} className="bg-white/10 p-2 rounded-full hover:text-red-500"><svg className="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" /></svg></button>
                            </div>
                            <div className="flex-1 flex items-center justify-center p-4 relative">
                                {modalImages.length > 1 && <button onClick={()=>setActiveImageIndex(p=>p===0?modalImages.length-1:p-1)} className="absolute left-4 p-3 bg-black/50 text-white rounded-full"><svg className="h-8 w-8" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M15 19l-7-7 7-7"/></svg></button>}
                                <img src={modalImages[activeImageIndex].data} alt="Evidencia Decodificada" className="max-h-full max-w-full object-contain" />
                                {modalImages.length > 1 && <button onClick={()=>setActiveImageIndex(p=>p===modalImages.length-1?0:p+1)} className="absolute right-4 p-3 bg-black/50 text-white rounded-full"><svg className="h-8 w-8" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 5l7 7-7 7"/></svg></button>}
                            </div>
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
