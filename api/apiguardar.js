// api/guardar.js
export default async function handler(req, res) {
    if (req.method !== 'POST') return res.status(405).send('Solo POST');

    const { CLIENT_ID, CLIENT_SECRET, TENANT_ID, SITE_ID, LIST_ID } = process.env;
    const data = req.body;

    try {
        // 1. Obtener Token de Microsoft
        const tokenResponse = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: new URLSearchParams({
                client_id: CLIENT_ID,
                client_secret: CLIENT_SECRET,
                scope: 'https://graph.microsoft.com/.default',
                grant_type: 'client_credentials'
            })
        });
        const { access_token } = await tokenResponse.json();

        // 2. Enviar datos a SharePoint (Crear item)
        const sharepointResponse = await fetch(`https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items`, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${access_token}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                fields: {
                    Title: data.id || "Checklist", // SharePoint siempre pide 'Title'
                    NombreSala: data.nombre,
                    HorasdeUso: data.horasUso,
                    HorasRestantes: data.horasRest,
                    EstadoHDMI: data.hdmi,
                    EstadoSwitch: data.switch,
                    EstadoEquiposPerifericos: data.equipos,
                    FechaUltimaRevision: data.fUlt,
                    ProximaRevision: data.fProx,
                    Responsable: data.responsable,
                    HoraReg: data.hora
                }
            })
        });

        if (sharepointResponse.ok) {
            res.status(200).json({ message: 'Guardado con éxito' });
        } else {
            res.status(500).json({ error: 'Error al escribir en SharePoint' });
        }
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
}