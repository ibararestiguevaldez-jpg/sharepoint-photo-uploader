const express = require('express');
const axios = require('axios');
const cors = require('cors');
require('dotenv').config();

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ limit: '50mb', extended: true }));

// Configuración de SharePoint
const SHAREPOINT_SITE = 'https://grupoelecmetal.sharepoint.com/sites/SeguimientodiariohornoD';
const SHAREPOINT_USER = process.env.SHAREPOINT_USER || 'iarestigue@cristalchile.cl';
const SHAREPOINT_PASS = process.env.SHAREPOINT_PASS || 'Cristal2029';
const PARENT_FOLDER = 'Fotos setup';

// Helper: Obtener Request Digest (necesario para POST/PUT)
async function getRequestDigest() {
  try {
    const url = `${SHAREPOINT_SITE}/_api/contextinfo`;
    const response = await axios.post(url, {}, {
      headers: {
        'Authorization': getBasicAuth(),
        'Accept': 'application/json',
        'Content-Type': 'application/json'
      }
    });
    return response.data.d.GetContextWebInformation.FormDigestValue;
  } catch (error) {
    console.error('Error obteniendo digest:', error.message);
    throw error;
  }
}

// Helper: Autenticación básica con SharePoint
function getBasicAuth() {
  const credentials = Buffer.from(`${SHAREPOINT_USER}:${SHAREPOINT_PASS}`).toString('base64');
  return `Basic ${credentials}`;
}

// Helper: Obtener ID de carpeta por nombre
async function getFolderId(folderName) {
  try {
    const url = `${SHAREPOINT_SITE}/_api/web/GetFolderByServerRelativeUrl('/sites/SeguimientodiariohornoD/Documentos%20compartidos/${PARENT_FOLDER}')`;
    const response = await axios.get(url, {
      headers: {
        'Authorization': getBasicAuth(),
        'Accept': 'application/json',
        'Content-Type': 'application/json'
      }
    });
    return response.data.UniqueId;
  } catch (error) {
    console.error('Error obteniendo ID de carpeta:', error.message);
    throw error;
  }
}

// Helper: Crear carpeta dentro de Fotos setup
async function createAlbumFolder(albumName) {
  try {
    const parentFolderId = await getFolderId(PARENT_FOLDER);
    const safeAlbumName = albumName.replace(/[<>:"/\\|?*]/g, '_').substring(0, 128);

    const url = `${SHAREPOINT_SITE}/_api/web/folders`;
    const response = await axios.post(
      url,
      {
        '__metadata': { type: 'SP.Folder' },
        'ServerRelativeUrl': `/sites/SeguimientodiariohornoD/Documentos compartidos/${PARENT_FOLDER}/${safeAlbumName}`
      },
      {
        headers: {
          'Authorization': getBasicAuth(),
          'Accept': 'application/json',
          'Content-Type': 'application/json',
          'X-RequestDigest': await getRequestDigest()
        }
      }
    );

    return {
      folderId: response.data.UniqueId,
      folderUrl: `/sites/SeguimientodiariohornoD/Documentos compartidos/${PARENT_FOLDER}/${safeAlbumName}`
    };
  } catch (error) {
    console.error('Error creando carpeta de álbum:', error.response?.data || error.message);
    throw error;
  }
}

// Helper: Subir archivo a SharePoint
async function uploadFileToSharePoint(base64Data, fileName, folderUrl) {
  try {
    const digest = await getRequestDigest();
    const buffer = Buffer.from(base64Data.split(',')[1], 'base64');

    const url = `${SHAREPOINT_SITE}/_api/web/GetFolderByServerRelativeUrl('${folderUrl}')/Files/add(url='${fileName}',overwrite=true)`;

    const response = await axios.post(url, buffer, {
      headers: {
        'Authorization': getBasicAuth(),
        'Accept': 'application/json',
        'Content-Type': 'application/octet-stream',
        'X-RequestDigest': digest
      }
    });

    return {
      success: true,
      fileName,
      serverRelativeUrl: response.data.d.ServerRelativeUrl
    };
  } catch (error) {
    console.error(`Error subiendo ${fileName}:`, error.response?.data || error.message);
    throw error;
  }
}

// Endpoint: Crear álbum y subir fotos
app.post('/api/upload-album', async (req, res) => {
  try {
    const { albumName, photos, albumData } = req.body;

    if (!albumName || !Array.isArray(photos) || photos.length === 0) {
      return res.status(400).json({
        error: 'Faltan datos: albumName y photos son requeridos'
      });
    }

    console.log(`📁 Creando carpeta: ${albumName}`);
    const { folderUrl } = await createAlbumFolder(albumName);

    console.log(`📸 Subiendo ${photos.length} fotos...`);
    const uploadResults = [];

    for (let i = 0; i < photos.length; i++) {
      const photo = photos[i];
      const fileName = `foto_${String(i + 1).padStart(2, '0')}.jpg`;

      try {
        const result = await uploadFileToSharePoint(photo.data, fileName, folderUrl);
        uploadResults.push(result);
        console.log(`✅ ${fileName} subida exitosamente`);
      } catch (error) {
        console.error(`❌ Error subiendo ${fileName}:`, error.message);
        uploadResults.push({
          fileName,
          success: false,
          error: error.message
        });
      }
    }

    // Guardar datos técnicos si existen
    if (albumData) {
      try {
        const jsonData = JSON.stringify(albumData, null, 2);
        const jsonBase64 = 'data:application/json;base64,' + Buffer.from(jsonData).toString('base64');
        await uploadFileToSharePoint(jsonBase64, 'datos_tecnicos.json', folderUrl);
        console.log('✅ Datos técnicos guardados');
      } catch (error) {
        console.error('⚠️ Error guardando datos técnicos:', error.message);
      }
    }

    const successCount = uploadResults.filter(r => r.success !== false).length;

    res.json({
      success: true,
      message: `Álbum "${albumName}" creado y ${successCount}/${photos.length} fotos subidas`,
      folderUrl: `${SHAREPOINT_SITE}/Documentos compartidos/${PARENT_FOLDER}/${albumName.replace(/[<>:"/\\|?*]/g, '_')}`,
      uploadResults,
      albumName,
      timestamp: new Date().toISOString()
    });

  } catch (error) {
    console.error('❌ Error en upload-album:', error.message);
    res.status(500).json({
      error: error.message,
      details: 'No se pudo crear el álbum en SharePoint'
    });
  }
});

// Endpoint: Health check
app.get('/api/health', (req, res) => {
  res.json({ status: 'OK', timestamp: new Date().toISOString() });
});

// Endpoint: Test de conexión a SharePoint
app.get('/api/test-connection', async (req, res) => {
  try {
    const digest = await getRequestDigest();
    res.json({
      success: true,
      message: 'Conexión a SharePoint exitosa'
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`🚀 Servidor ejecutándose en puerto ${PORT}`);
});

module.exports = app;