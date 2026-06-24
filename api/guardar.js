export default async function handler(req, res) {
    // Solo permitir peticiones POST
    if (req.method !== 'POST') {
        return res.status(405).json({ error: "Método no permitido" });
    }

    // Obtener variables de entorno
    const { CLIENT_ID, CLIENT_SECRET, TENANT_ID, SITE_ID, LIST_ID } = process.env;

    // Depuración: Verificar que las variables existen
    if (!CLIENT_ID || !CLIENT_SECRET || !TENANT_ID || !SITE_ID || !LIST_ID) {
        console.error("Error: Faltan variables de entorno en Vercel");
        return res.status(500).json({ error: "Faltan variables de configuración en el servidor" });
    }

    const data = req.body;
    console.log("Datos recibidos de la App:", JSON.stringify(data));

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

        const tokenData = await tokenResponse.json();
        
        if (!tokenData.access_token) {
            console.error("Error al obtener token:", tokenData);
            return res.status(500).json({ error: "No se pudo obtener token de Microsoft" });
        }

        // 2. Enviar datos a SharePoint
        const result = await fetch(`https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items`, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${tokenData.access_token}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                fields: {
                    Title: data.id || "Checklist",
                    IDSala: data.id,
                    NombreSala: data.nombre,
                    HorasdeUso: data.horasUso,
                    EstadoHDMI: data.hdmi,
                    EstadoSwitch: data.switch,
                    EstadoequiposP: data.equipos,
                    FechaUltimaRevisin: data.fUlt,
                    ProximaRevisin: data.fProx,
                    Responsable: data.responsable,
                    FechaHora: data.hora
                }
            })
        });

        const responseData = await result.json();

        if (result.ok) {
            console.log("Guardado exitoso en SharePoint");
            res.status(200).json({ success: true });
        } else {
            console.error("Error de SharePoint:", JSON.stringify(responseData));
            res.status(500).json({ error: responseData });
        }
    } catch (e) {
        console.error("Error crítico en el servidor:", e.message);
        res.status(500).json({ error: e.message });
    }
}
