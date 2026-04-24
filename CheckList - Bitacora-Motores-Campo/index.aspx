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
    </style>
</head>
<body class="antialiased text-gray-800">
    <div id="root"></div>

    <script type="text/babel">
        const { useState, useEffect } = React;

        const SP_CONFIG = {
            siteUrl: "https://glencore.sharepoint.com/sites/co-lmn-sgia/checklist",
            listTitle: "Bitacora-Motores-Campo" 
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

        function App() {
            const [formData, setFormData] = useState(initialFormState);
            const [items, setItems] = useState([]);
            const [loading, setLoading] = useState(false);
            const [error, setError] = useState(null);
            const [editingId, setEditingId] = useState(null);

            useEffect(() => { fetchItems(); }, []);

            const getRequestDigest = async () => {
                const response = await fetch(`${SP_CONFIG.siteUrl}/_api/contextinfo`, {
                    method: 'POST',
                    headers: { "Accept": "application/json;odata=verbose" }
                });
                const data = await response.json();
                return data.d.GetContextWebInformation.FormDigestValue;
            };

            const getEntityType = async () => {
                const response = await fetch(`${SP_CONFIG.siteUrl}/_api/web/lists/getbytitle('${SP_CONFIG.listTitle}')?$select=ListItemEntityTypeFullName`, {
                    headers: { "Accept": "application/json;odata=verbose" }
                });
                const data = await response.json();
                return data.d.ListItemEntityTypeFullName;
            };

            const fetchItems = async () => {
                setLoading(true);
                try {
                    const response = await fetch(`${SP_CONFIG.siteUrl}/_api/web/lists/getbytitle('${SP_CONFIG.listTitle}')/items?$top=50&$orderby=Created desc`, {
                        headers: { "Accept": "application/json;odata=verbose" }
                    });
                    if (!response.ok) throw new Error("Error en la red al consultar");
                    const data = await response.json();
                    setItems(data.d.results);
                } catch (err) {
                    setError("Fallo al cargar registros. Revisa tu conexi\u00F3n a SharePoint.");
                } finally {
                    setLoading(false);
                }
            };

            const handleChange = (e) => {
                setFormData({ ...formData, [e.target.name]: e.target.value });
            };

            const handleSubmit = async (e) => {
                e.preventDefault();
                setLoading(true);
                setError(null);

                try {
                    const digest = await getRequestDigest();
                    const entityType = await getEntityType();
                    
                    let formattedDate = null;
                    if (formData.Fecha) {
                        formattedDate = new Date(`${formData.Fecha}T12:00:00Z`).toISOString();
                    }

                    const itemPayload = {
                        "__metadata": { "type": entityType },
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

                    if (!response.ok) {
                        const errorData = await response.json();
                        throw new Error(errorData.error.message.value);
                    }
                    
                    setFormData(initialFormState);
                    setEditingId(null);
                    await fetchItems();
                } catch (err) {
                    setError(`Error al guardar: ${err.message}`);
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
                window.scrollTo({ top: 0, behavior: 'smooth' });
            };

            const glassCard = "bg-white/70 backdrop-blur-xl border border-white/40 shadow-2xl rounded-2xl";
            const inputClass = "block w-full rounded-xl bg-white/60 border-white/50 shadow-inner focus:bg-white focus:border-cerrejon-orange focus:ring-2 focus:ring-cerrejon-orange/50 transition-all duration-300 p-3 text-sm font-medium outline-none";
            const labelClass = "block text-xs font-bold text-gray-700 mb-2 uppercase tracking-widest";

            return (
                <div className="max-w-7xl mx-auto py-10 px-4 sm:px-6 lg:px-8 min-h-screen flex flex-col gap-8">
                    
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

                    <div className="grid grid-cols-1 lg:grid-cols-12 gap-8 items-start">
                        
                        <div className={`${glassCard} p-8 lg:col-span-7 transition-all duration-500`}>
                            <div className="mb-6 flex items-center justify-between">
                                <h2 className="text-xl font-extrabold text-gray-800 tracking-tight">
                                    {editingId ? "Actualizar Registro" : "Nuevo Registro"}
                                </h2>
                                {editingId && <span className="px-3 py-1 bg-cerrejon-gold/20 text-cerrejon-orange text-xs font-bold rounded-full uppercase">Modo Edici&oacute;n</span>}
                            </div>

                            <form onSubmit={handleSubmit} className="space-y-6">
                                
                                <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
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

                                <div>
                                    <label className={labelClass}>Diagn&oacute;stico Inicial</label>
                                    <textarea name="Diagnostico" value={formData.Diagnostico} onChange={handleChange} rows="2" className={`${inputClass} resize-none`} placeholder="Describe brevemente la falla..."></textarea>
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
                                
                                <div>
                                    <label className={labelClass}>Detalle del Backlog</label>
                                    <textarea name="BacklogRegistrado" value={formData.BacklogRegistrado} onChange={handleChange} rows="2" className={`${inputClass} resize-none`} placeholder="Especificaciones del backlog..."></textarea>
                                </div>

                                <div>
                                    <label className={labelClass}>Trabajos Pendientes</label>
                                    <textarea name="Pendientes" value={formData.Pendientes} onChange={handleChange} rows="2" className={`${inputClass} resize-none`} placeholder="Repuestos, firmas, aprobaciones..."></textarea>
                                </div>

                                <div className="flex justify-end gap-3 pt-4 border-t border-white/30">
                                    {editingId && (
                                        <button type="button" onClick={() => {setEditingId(null); setFormData(initialFormState);}} className="px-5 py-3 bg-white/50 text-gray-700 font-bold rounded-xl hover:bg-white/80 transition-colors">
                                            Cancelar
                                        </button>
                                    )}
                                    <button type="submit" disabled={loading} className="px-8 py-3 bg-gradient-to-r from-cerrejon-orange to-red-600 text-white font-bold rounded-xl shadow-lg hover:shadow-xl hover:scale-[1.02] transition-all disabled:opacity-50 disabled:scale-100">
                                        {loading ? "Procesando Datos..." : (editingId ? "Guardar Cambios" : "Crear Registro")}
                                    </button>
                                </div>
                            </form>
                        </div>

                        <div className={`${glassCard} flex flex-col lg:col-span-5 h-[800px] overflow-hidden`}>
                            <div className="bg-gray-900/80 backdrop-blur-md px-6 py-5 flex justify-between items-center">
                                <h2 className="text-sm font-bold text-white uppercase tracking-widest">Registros Recientes</h2>
                                <span className="bg-cerrejon-orange/20 text-cerrejon-gold text-xs px-2 py-1 rounded-md font-bold">{items.length} totales</span>
                            </div>
                            
                            <div className="overflow-y-auto flex-grow p-4 space-y-3 no-scrollbar">
                                {items.length === 0 ? (
                                    <div className="text-center p-8 text-gray-500 font-medium">No hay bit&aacute;coras registradas.</div>
                                ) : (
                                    items.map((item) => {
                                        const isAvailable = item.EquipoDisponible === 'SI';
                                        return (
                                            <div key={item.Id} className="bg-white/60 hover:bg-white/90 border border-white/50 p-4 rounded-xl shadow-sm hover:shadow-md transition-all group relative">
                                                <div className="flex justify-between items-start mb-2">
                                                    <div className="flex items-center gap-2">
                                                        <span className="text-xs font-black text-cerrejon-orange bg-orange-100 px-2 py-1 rounded">OT-{item.OT || item.Title}</span>
                                                        <span className="text-xs font-bold text-gray-500">{item.Fecha ? item.Fecha.split('T')[0] : "Sin fecha"}</span>
                                                    </div>
                                                    <button onClick={() => handleEdit(item)} className="text-gray-400 hover:text-cerrejon-orange transition-colors">
                                                        <svg xmlns="http://www.w3.org/2000/svg" className="h-5 w-5" viewBox="0 0 20 20" fill="currentColor">
                                                            <path d="M13.586 3.586a2 2 0 112.828 2.828l-.793.793-2.828-2.828.793-.793zM11.379 5.793L3 14.172V17h2.828l8.38-8.379-2.83-2.828z" />
                                                        </svg>
                                                    </button>
                                                </div>
                                                
                                                <div className="mb-3">
                                                    <h3 className="font-black text-gray-800 text-lg">{item.Equipo}</h3>
                                                    <p className="text-xs text-gray-600 font-medium mt-1 truncate" title={item.Diagnostico}>{item.Diagnostico || "Sin diagn\u00F3stico registrado"}</p>
                                                </div>

                                                <div className="flex items-center justify-between mt-4 pt-3 border-t border-gray-200/50">
                                                    <div className="flex items-center gap-2">
                                                        <div className={`w-2 h-2 rounded-full ${isAvailable ? 'bg-green-500' : 'bg-red-500'}`}></div>
                                                        <span className="text-xs font-bold text-gray-600">{isAvailable ? 'Operativo' : 'Fuera de Servicio'}</span>
                                                    </div>
                                                    <span className="text-xs text-gray-500 font-medium truncate max-w-[120px]">{item.Tecnicos}</span>
                                                </div>
                                            </div>
                                        );
                                    })
                                )}
                            </div>
                        </div>

                    </div>
                </div>
            );
        }

        const root = ReactDOM.createRoot(document.getElementById('root'));
        root.render(<App />);
    </script>
</body>
</html>