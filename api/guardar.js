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
        const tokenData = await tokenResponse.json();

        // AQUÍ ESTÁ EL CAMBIO: Solo enviamos el 'Title' para confirmar conexión
        const result = await fetch(`https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items`, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer ${tokenData.access_token}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                fields: {
                    Title: data.id || "Checklist" 
                    // Si esto funciona, iremos descomentando los otros campos uno a uno
                }
            })
        });

        const responseData = await result.json();

        if (result.ok) {
            res.status(200).json({ success: true });
        } else {
            console.error("Error de SharePoint:", JSON.stringify(responseData));
            res.status(500).json({ error: responseData });
        }
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
}
