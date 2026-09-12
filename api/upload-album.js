const axios = require('axios');

const SHAREPOINT_SITE = 'https://grupoelecmetal.sharepoint.com/sites/SeguimientodiariohornoD';
const SHAREPOINT_USER = process.env.SHAREPOINT_USER || 'iarestigue@cristalchile.cl';
const SHAREPOINT_PASS = process.env.SHAREPOINT_PASS || 'Cristal2029';
const PARENT_FOLDER = 'Fotos setup';

function getBasicAuth() {
  const credentials = Buffer.from(`${SHAREPOINT_USER}:${SHAREPOINT_PASS}`).toString('base64');
  return `Basic ${credentials}`;
}

async function getRequestDigest() {
  const url = `${SHAREPOINT_SITE}/_api/contextinfo`;
  const response = await axios.post(url, {}, {
    headers: {
      'Authorization': getBasicAuth(),
      'Accept': 'application/json',
      'Content-Type': 'application/json'
    }
  });
  return response.data.d.GetContextWebInformation.FormDigestValue;
}

async function createAlbumFolder(albumName) {
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
}

async function uploadFileToSharePoint(base64Data, fileName, folderUrl) {
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
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Credentials', 'true');
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS,PATCH,DELETE,POST,PUT');
  res.setHeader('Access-Control-Allow-Headers', 'X-CSRF-Token, X-Requested-With, Accept, Accept-Version, Content-Length, Content-MD5, Content-Type, Date, X-Api-Version');

  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  try {
    const { albumName, photos, albumData } = req.body;

    if (!albumName || !Array.isArray(photos) || photos.length === 0) {
      return res.status(400).json({ error: 'Faltan datos: albumName y photos son requeridos' });
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
      message: `Álbum "${albumName}" creado y fotos subidas`,
      folderUrl: `${SHAREPOINT_SITE}/Documentos compartidos/${PARENT_FOLDER}/${albumName.replace(/[<>:"/\\|?*]/g, '_')}`,
      uploadResults,
      albumName,
      timestamp: new Date().toISOString()
    });
  } catch (error) {
    res.status(500).json({ error: error.message, details: 'No se pudo crear el álbum en SharePoint' });
  }
}
