const express = require('express');
const axios = require('axios');
const cors = require('cors');
require('dotenv').config();

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ limit: '50mb', extended: true }));

const SHAREPOINT_SITE = 'https://grupoelecmetal.sharepoint.com/sites/SeguimientodiariohornoD';
const SHAREPOINT_USER = process.env.SHAREPOINT_USER || 'iarestigue@cristalchile.cl';
const SHAREPOINT_PASS = process.env.SHAREPOINT_PASS || 'Cristal2029';
const PARENT_FOLDER = 'Fotos setup';

function getBasicAuth() {
  const credentials = Buffer.from(`${SHAREPOINT_USER}:${SHAREPOINT_PASS}`).toString('base64');
  return `Basic ${credentials}`;
}

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

async function createAlbumFolder(albumName) {
  try {
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
    console.error('Error creando carpeta:', error.message);
    throw error;
  }
}

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

    return { success: true, fileName, serverRelativeUrl: response.data.d.ServerRelativeUrl };
  } catch (error) {
    console.error(`Error subiendo ${fileName}:`, error.message);
    throw error;
  }
}

app.post('/api/upload-album', async (req, res) => {
  try {
    const { albumName, photos, albumData } = req.body;

    if (!albumName || !Array.isArray(photos) || photos.length === 0) {
      return res.status(400).json({ error: 'Faltan datos' });
    }

    const { folderUrl } = await createAlbumFolder(albumName);
    const uploadResults = [];

    for (let i = 0; i < photos.length; i++) {
      const photo = photos[i];
      const fileName = `foto_${String(i + 1).padStart(2, '0')}.jpg`;
      try {
        const result = await uploadFileToSharePoint(photo.data, fileName, folderUrl);
        uploadResults.push(result);
      } catch (error) {
        uploadResults.push({ fileName, success: false, error: error.message });
      }
    }

    res.json({
      success: true,
      message: `Álbum "${albumName}" creado`,
      folderUrl: `${SHAREPOINT_SITE}/Documentos compartidos/${PARENT_FOLDER}/${albumName.replace(/[<>:"/\\|?*]/g, '_')}`,
      uploadResults,
      albumName,
      timestamp: new Date().toISOString()
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.get('/api/health', (req, res) => {
  res.json({ status: 'OK', timestamp: new Date().toISOString() });
});

module.exports = app;
