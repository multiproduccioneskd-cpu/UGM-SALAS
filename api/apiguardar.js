export default async function handler(req, res) {
    if (req.method !== 'POST') return res.status(405).end();
    
    const { CLIENT_ID, CLIENT_SECRET, TENANT_ID, SITE_ID, LIST_ID } = process.env;
    const data = req.body;

    try {
        // 1. Obtener Token
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
        if (!tokenData.access_token) throw new Error("No se pudo obtener el token");

        // 2. Enviar datos a SharePoint
        // Usamos los nombres más probables. Si falla, revisa el error en los logs de Vercel.
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
            res.status(200).json({ success: true, data: responseData });
        } else {
            // AQUÍ VERÁS EL ERROR EXACTO EN LOS LOGS DE VERCEL SI FALLA
            console.error("Error de Microsoft:", JSON.stringify(responseData));
            res.status(500).json({ error: responseData });
        }
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
}
