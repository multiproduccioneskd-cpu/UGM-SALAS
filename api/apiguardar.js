export default async function handler(req, res) {
    if (req.method !== 'POST') return res.status(405).end();
    
    const { CLIENT_ID, CLIENT_SECRET, TENANT_ID, SITE_ID, LIST_ID } = process.env;
    const data = req.body;

    try {
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

        const result = await fetch(`https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items`, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${access_token}`,
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

        if (result.ok) {
            res.status(200).json({ success: true });
        } else {
            const error = await result.json();
            res.status(500).json({ error });
        }
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
}
