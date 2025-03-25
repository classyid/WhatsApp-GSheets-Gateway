/**
 * WhatsApp Gateway Integration for Google Sheets
 * Menggunakan API dari https://m-pedia.co.id/
 */

// Konstanta untuk API
const API_KEY = "<apikey>"; // API Key dari dokumen
const DEFAULT_SENDER = "<sender>"; // Sender default dari dokumen
const BASE_URL = "https://url-mpedia";

/**
 * Membuat menu di toolbar Google Spreadsheet saat file dibuka
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('WhatsApp Gateway')
    .addItem('Kirim Pesan', 'showSendMessageDialog')
    .addItem('Kirim Media', 'showSendMediaDialog')
    .addItem('Kirim Stiker', 'showSendStickerDialog')
    .addItem('Kirim vCard', 'showSendVCardDialog')
    .addItem('Cek Nomor', 'showCheckNumberDialog')
    .addItem('Informasi Device', 'showDeviceInfoDialog')
    .addItem('Generate QR', 'showGenerateQRDialog')
    .addItem('Logout Device', 'showLogoutDeviceDialog')
    .addSeparator()
    .addItem('Kirim Pesan Massal dari Sheet', 'showBulkMessageDialog')
    .addSeparator()
    .addItem('Pengaturan', 'showSettingsDialog')
    .addToUi();
}

/**
 * Menyimpan pengaturan di Properties
 */
function saveSettings(apiKey, sender) {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperties({
    'apiKey': apiKey,
    'sender': sender
  });
}

/**
 * Mendapatkan pengaturan dari Properties
 */
function getSettings() {
  const userProperties = PropertiesService.getUserProperties();
  const apiKey = userProperties.getProperty('apiKey') || API_KEY;
  const sender = userProperties.getProperty('sender') || DEFAULT_SENDER;
  
  return {
    apiKey: apiKey,
    sender: sender
  };
}

/**
 * Menampilkan dialog pengaturan
 */
function showSettingsDialog() {
  const settings = getSettings();
  const htmlOutput = HtmlService
    .createHtmlOutput(`
      <form>
        <div style="margin-bottom: 10px;">
          <label for="apiKey">API Key:</label><br>
          <input type="text" id="apiKey" name="apiKey" value="${settings.apiKey}" style="width: 100%;">
        </div>
        <div style="margin-bottom: 10px;">
          <label for="sender">Sender Number:</label><br>
          <input type="text" id="sender" name="sender" value="${settings.sender}" style="width: 100%;">
        </div>
        <div style="text-align: center; margin-top: 20px;">
          <input type="button" value="Simpan" onclick="saveSettings()" style="padding: 5px 15px;">
        </div>
      </form>
      <script>
        function saveSettings() {
          const apiKey = document.getElementById('apiKey').value;
          const sender = document.getElementById('sender').value;
          google.script.run.withSuccessHandler(closeDialog).saveSettings(apiKey, sender);
        }
        
        function closeDialog() {
          google.script.host.close();
        }
      </script>
    `)
    .setWidth(400)
    .setHeight(200)
    .setTitle('Pengaturan WhatsApp Gateway');
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Pengaturan WhatsApp Gateway');
}

/**
 * Menampilkan dialog kirim pesan
 */
function showSendMessageDialog() {
  const settings = getSettings();
  const htmlOutput = HtmlService
    .createHtmlOutput(`
      <form>
        <div style="margin-bottom: 10px;">
          <label for="number">Nomor Tujuan:</label><br>
          <input type="text" id="number" name="number" placeholder="628xxxxxxxxxx" style="width: 100%;">
        </div>
        <div style="margin-bottom: 10px;">
          <label for="message">Pesan:</label><br>
          <textarea id="message" name="message" rows="5" style="width: 100%;"></textarea>
        </div>
        <div style="text-align: center; margin-top: 20px;">
          <input type="button" value="Kirim" onclick="sendMessage()" style="padding: 5px 15px;">
        </div>
      </form>
      <div id="result" style="margin-top: 10px; text-align: center;"></div>
      <script>
        function sendMessage() {
          const number = document.getElementById('number').value;
          const message = document.getElementById('message').value;
          
          if (!number || !message) {
            document.getElementById('result').innerHTML = 'Nomor dan pesan harus diisi!';
            return;
          }
          
          document.getElementById('result').innerHTML = 'Mengirim...';
          google.script.run
            .withSuccessHandler(showResult)
            .withFailureHandler(showError)
            .sendMessage(number, message);
        }
        
        function showResult(result) {
          document.getElementById('result').innerHTML = result;
        }
        
        function showError(error) {
          document.getElementById('result').innerHTML = 'Error: ' + error.message;
        }
      </script>
    `)
    .setWidth(400)
    .setHeight(350)
    .setTitle('Kirim Pesan WhatsApp');
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Kirim Pesan WhatsApp');
}

/**
 * Fungsi untuk mengirim pesan WhatsApp
 */
function sendMessage(number, message) {
  const settings = getSettings();
  
  const payload = {
    api_key: settings.apiKey,
    sender: settings.sender,
    number: number,
    message: message
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  };
  
  try {
    const response = UrlFetchApp.fetch(`${BASE_URL}/send-message`, options);
    const result = JSON.parse(response.getContentText());
    
    if (result.status) {
      return 'Pesan berhasil dikirim!';
    } else {
      return 'Gagal mengirim pesan: ' + result.message;
    }
  } catch (error) {
    return 'Error: ' + error.toString();
  }
}

/**
 * Menampilkan dialog kirim media
 */
function showSendMediaDialog() {
  const htmlOutput = HtmlService
    .createHtmlOutput(`
      <form>
        <div style="margin-bottom: 10px;">
          <label for="number">Nomor Tujuan:</label><br>
          <input type="text" id="number" name="number" placeholder="628xxxxxxxxxx" style="width: 100%;">
        </div>
        <div style="margin-bottom: 10px;">
          <label for="mediaType">Tipe Media:</label><br>
          <select id="mediaType" name="mediaType" style="width: 100%;">
            <option value="image">Gambar</option>
            <option value="video">Video</option>
            <option value="audio">Audio</option>
            <option value="document">Dokumen</option>
          </select>
        </div>
        <div style="margin-bottom: 10px;">
          <label for="url">URL Media (direct link):</label><br>
          <input type="text" id="url" name="url" style="width: 100%;">
        </div>
        <div style="margin-bottom: 10px;">
          <label for="caption">Caption (opsional):</label><br>
          <textarea id="caption" name="caption" rows="3" style="width: 100%;"></textarea>
        </div>
        <div style="text-align: center; margin-top: 20px;">
          <input type="button" value="Kirim" onclick="sendMedia()" style="padding: 5px 15px;">
        </div>
      </form>
      <div id="result" style="margin-top: 10px; text-align: center;"></div>
      <script>
        function sendMedia() {
          const number = document.getElementById('number').value;
          const mediaType = document.getElementById('mediaType').value;
          const url = document.getElementById('url').value;
          const caption = document.getElementById('caption').value;
          
          if (!number || !url) {
            document.getElementById('result').innerHTML = 'Nomor dan URL harus diisi!';
            return;
          }
          
          document.getElementById('result').innerHTML = 'Mengirim...';
          google.script.run
            .withSuccessHandler(showResult)
            .withFailureHandler(showError)
            .sendMedia(number, mediaType, url, caption);
        }
        
        function showResult(result) {
          document.getElementById('result').innerHTML = result;
        }
        
        function showError(error) {
          document.getElementById('result').innerHTML = 'Error: ' + error.message;
        }
      </script>
    `)
    .setWidth(400)
    .setHeight(400)
    .setTitle('Kirim Media WhatsApp');
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Kirim Media WhatsApp');
}

/**
 * Fungsi untuk mengirim media WhatsApp
 */
function sendMedia(number, mediaType, url, caption) {
  const settings = getSettings();
  
  const payload = {
    api_key: settings.apiKey,
    sender: settings.sender,
    number: number,
    media_type: mediaType,
    url: url
  };
  
  if (caption) {
    payload.caption = caption;
  }
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  };
  
  try {
    const response = UrlFetchApp.fetch(`${BASE_URL}/send-media`, options);
    const result = JSON.parse(response.getContentText());
    
    if (result.status) {
      return 'Media berhasil dikirim!';
    } else {
      return 'Gagal mengirim media: ' + result.message;
    }
  } catch (error) {
    return 'Error: ' + error.toString();
  }
}

/**
 * Menampilkan dialog kirim stiker
 */
function showSendStickerDialog() {
  const htmlOutput = HtmlService
    .createHtmlOutput(`
      <form>
        <div style="margin-bottom: 10px;">
          <label for="number">Nomor Tujuan:</label><br>
          <input type="text" id="number" name="number" placeholder="628xxxxxxxxxx" style="width: 100%;">
        </div>
        <div style="margin-bottom: 10px;">
          <label for="url">URL Stiker (direct link):</label><br>
          <input type="text" id="url" name="url" style="width: 100%;">
        </div>
        <div style="text-align: center; margin-top: 20px;">
          <input type="button" value="Kirim" onclick="sendSticker()" style="padding: 5px 15px;">
        </div>
      </form>
      <div id="result" style="margin-top: 10px; text-align: center;"></div>
      <script>
        function sendSticker() {
          const number = document.getElementById('number').value;
          const url = document.getElementById('url').value;
          
          if (!number || !url) {
            document.getElementById('result').innerHTML = 'Nomor dan URL harus diisi!';
            return;
          }
          
          document.getElementById('result').innerHTML = 'Mengirim...';
          google.script.run
            .withSuccessHandler(showResult)
            .withFailureHandler(showError)
            .sendSticker(number, url);
        }
        
        function showResult(result) {
          document.getElementById('result').innerHTML = result;
        }
        
        function showError(error) {
          document.getElementById('result').innerHTML = 'Error: ' + error.message;
        }
      </script>
    `)
    .setWidth(400)
    .setHeight(250)
    .setTitle('Kirim Stiker WhatsApp');
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Kirim Stiker WhatsApp');
}

/**
 * Fungsi untuk mengirim stiker WhatsApp
 */
function sendSticker(number, url) {
  const settings = getSettings();
  
  const payload = {
    api_key: settings.apiKey,
    sender: settings.sender,
    number: number,
    url: url
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  };
  
  try {
    const response = UrlFetchApp.fetch(`${BASE_URL}/send-sticker`, options);
    const result = JSON.parse(response.getContentText());
    
    if (result.status) {
      return 'Stiker berhasil dikirim!';
    } else {
      return 'Gagal mengirim stiker: ' + result.message;
    }
  } catch (error) {
    return 'Error: ' + error.toString();
  }
}

/**
 * Menampilkan dialog kirim vCard
 */
function showSendVCardDialog() {
  const htmlOutput = HtmlService
    .createHtmlOutput(`
      <form>
        <div style="margin-bottom: 10px;">
          <label for="number">Nomor Tujuan:</label><br>
          <input type="text" id="number" name="number" placeholder="628xxxxxxxxxx" style="width: 100%;">
        </div>
        <div style="margin-bottom: 10px;">
          <label for="name">Nama Kontak:</label><br>
          <input type="text" id="name" name="name" style="width: 100%;">
        </div>
        <div style="margin-bottom: 10px;">
          <label for="phone">Nomor Telepon Kontak:</label><br>
          <input type="text" id="phone" name="phone" placeholder="628xxxxxxxxxx" style="width: 100%;">
        </div>
        <div style="text-align: center; margin-top: 20px;">
          <input type="button" value="Kirim" onclick="sendVCard()" style="padding: 5px 15px;">
        </div>
      </form>
      <div id="result" style="margin-top: 10px; text-align: center;"></div>
      <script>
        function sendVCard() {
          const number = document.getElementById('number').value;
          const name = document.getElementById('name').value;
          const phone = document.getElementById('phone').value;
          
          if (!number || !name || !phone) {
            document.getElementById('result').innerHTML = 'Semua field harus diisi!';
            return;
          }
          
          document.getElementById('result').innerHTML = 'Mengirim...';
          google.script.run
            .withSuccessHandler(showResult)
            .withFailureHandler(showError)
            .sendVCard(number, name, phone);
        }
        
        function showResult(result) {
          document.getElementById('result').innerHTML = result;
        }
        
        function showError(error) {
          document.getElementById('result').innerHTML = 'Error: ' + error.message;
        }
      </script>
    `)
    .setWidth(400)
    .setHeight(300)
    .setTitle('Kirim vCard WhatsApp');
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Kirim vCard WhatsApp');
}

/**
 * Fungsi untuk mengirim vCard WhatsApp
 */
function sendVCard(number, name, phone) {
  const settings = getSettings();
  
  const payload = {
    api_key: settings.apiKey,
    sender: settings.sender,
    number: number,
    name: name,
    phone: phone
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  };
  
  try {
    const response = UrlFetchApp.fetch(`${BASE_URL}/send-vcard`, options);
    const result = JSON.parse(response.getContentText());
    
    if (result.status) {
      return 'vCard berhasil dikirim!';
    } else {
      return 'Gagal mengirim vCard: ' + result.message;
    }
  } catch (error) {
    return 'Error: ' + error.toString();
  }
}

/**
 * Menampilkan dialog cek nomor
 */
function showCheckNumberDialog() {
  const htmlOutput = HtmlService
    .createHtmlOutput(`
      <form>
        <div style="margin-bottom: 10px;">
          <label for="number">Nomor yang akan dicek:</label><br>
          <input type="text" id="number" name="number" placeholder="628xxxxxxxxxx" style="width: 100%;">
        </div>
        <div style="text-align: center; margin-top: 20px;">
          <input type="button" value="Cek" onclick="checkNumber()" style="padding: 5px 15px;">
        </div>
      </form>
      <div id="result" style="margin-top: 10px; text-align: center;"></div>
      <script>
        function checkNumber() {
          const number = document.getElementById('number').value;
          
          if (!number) {
            document.getElementById('result').innerHTML = 'Nomor harus diisi!';
            return;
          }
          
          document.getElementById('result').innerHTML = 'Mengecek...';
          google.script.run
            .withSuccessHandler(showResult)
            .withFailureHandler(showError)
            .checkNumber(number);
        }
        
        function showResult(result) {
          document.getElementById('result').innerHTML = result;
        }
        
        function showError(error) {
          document.getElementById('result').innerHTML = 'Error: ' + error.message;
        }
      </script>
    `)
    .setWidth(400)
    .setHeight(200)
    .setTitle('Cek Nomor WhatsApp');
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Cek Nomor WhatsApp');
}

/**
 * Fungsi untuk memeriksa nomor WhatsApp
 */
function checkNumber(number) {
  const settings = getSettings();
  
  const payload = {
    api_key: settings.apiKey,
    sender: settings.sender,
    number: number
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  };
  
  try {
    const response = UrlFetchApp.fetch(`${BASE_URL}/check-number`, options);
    const result = JSON.parse(response.getContentText());
    
    if (result.status) {
      if (result.msg.exists) {
        return `Nomor ${number} terdaftar di WhatsApp dengan JID: ${result.msg.jid}`;
      } else {
        return `Nomor ${number} tidak terdaftar di WhatsApp`;
      }
    } else {
      return 'Gagal memeriksa nomor: ' + result.message;
    }
  } catch (error) {
    return 'Error: ' + error.toString();
  }
}

/**
 * Menampilkan dialog informasi device
 */
function showDeviceInfoDialog() {
  const htmlOutput = HtmlService
    .createHtmlOutput(`
      <form>
        <div style="margin-bottom: 10px;">
          <label for="number">Nomor Device:</label><br>
          <input type="text" id="number" name="number" placeholder="628xxxxxxxxxx" style="width: 100%;">
        </div>
        <div style="text-align: center; margin-top: 20px;">
          <input type="button" value="Cek" onclick="getDeviceInfo()" style="padding: 5px 15px;">
        </div>
      </form>
      <div id="result" style="margin-top: 10px;"></div>
      <script>
        function getDeviceInfo() {
          const number = document.getElementById('number').value;
          
          if (!number) {
            document.getElementById('result').innerHTML = 'Nomor harus diisi!';
            return;
          }
          
          document.getElementById('result').innerHTML = 'Mengambil informasi...';
          google.script.run
            .withSuccessHandler(showResult)
            .withFailureHandler(showError)
            .getDeviceInfo(number);
        }
        
        function showResult(result) {
          document.getElementById('result').innerHTML = result;
        }
        
        function showError(error) {
          document.getElementById('result').innerHTML = 'Error: ' + error.message;
        }
      </script>
    `)
    .setWidth(400)
    .setHeight(300)
    .setTitle('Informasi Device WhatsApp');
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Informasi Device WhatsApp');
}

/**
 * Fungsi untuk mendapatkan informasi device
 */
function getDeviceInfo(number) {
  const settings = getSettings();
  
  const payload = {
    api_key: settings.apiKey,
    number: number
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  };
  
  try {
    const response = UrlFetchApp.fetch(`${BASE_URL}/info-device`, options);
    const result = JSON.parse(response.getContentText());
    
    if (result.status && result.info.length > 0) {
      const device = result.info[0];
      
      return `
        <div style="font-family: Arial; font-size: 14px;">
          <p><strong>ID:</strong> ${device.id}</p>
          <p><strong>Nomor:</strong> ${device.body}</p>
          <p><strong>Status:</strong> ${device.status}</p>
          <p><strong>Pesan Terkirim:</strong> ${device.message_sent}</p>
          <p><strong>Dibuat pada:</strong> ${device.created_at}</p>
        </div>
      `;
    } else if (result.status && result.info.length === 0) {
      return 'Tidak ada informasi device untuk nomor tersebut';
    } else {
      return 'Gagal mendapatkan informasi device: ' + result.message;
    }
  } catch (error) {
    return 'Error: ' + error.toString();
  }
}

/**
 * Menampilkan dialog generate QR
 */
function showGenerateQRDialog() {
  const htmlOutput = HtmlService
    .createHtmlOutput(`
      <form>
        <div style="margin-bottom: 10px;">
          <label for="device">Nomor Device:</label><br>
          <input type="text" id="device" name="device" placeholder="628xxxxxxxxxx" style="width: 100%;">
        </div>
        <div style="margin-bottom: 10px;">
          <label for="force">Buat baru jika tidak ada:</label><br>
          <input type="checkbox" id="force" name="force" value="true">
        </div>
        <div style="text-align: center; margin-top: 20px;">
          <input type="button" value="Generate QR" onclick="generateQR()" style="padding: 5px 15px;">
        </div>
      </form>
      <div id="result" style="margin-top: 10px; text-align: center;"></div>
      <script>
        function generateQR() {
          const device = document.getElementById('device').value;
          const force = document.getElementById('force').checked;
          
          if (!device) {
            document.getElementById('result').innerHTML = 'Nomor device harus diisi!';
            return;
          }
          
          document.getElementById('result').innerHTML = 'Generating QR...';
          google.script.run
            .withSuccessHandler(showResult)
            .withFailureHandler(showError)
            .generateQR(device, force);
        }
        
        function showResult(result) {
          document.getElementById('result').innerHTML = result;
        }
        
        function showError(error) {
          document.getElementById('result').innerHTML = 'Error: ' + error.message;
        }
      </script>
    `)
    .setWidth(400)
    .setHeight(300)
    .setTitle('Generate QR WhatsApp');
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Generate QR WhatsApp');
}

/**
 * Fungsi untuk generate QR
 */
function generateQR(device, force) {
  const settings = getSettings();
  
  const payload = {
    api_key: settings.apiKey,
    device: device
  };
  
  if (force) {
    payload.force = true;
  }
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  };
  
  try {
    const response = UrlFetchApp.fetch(`${BASE_URL}/generate-qr`, options);
    const result = JSON.parse(response.getContentText());
    
    // Log respon lengkap untuk debugging
    Logger.log("Respon QR API: " + JSON.stringify(result));
    
    if (result.status) {
      // Cek berbagai kemungkinan nama field untuk QR code
      const qrValue = result.qrcode || result.qr || result.data?.qrcode || result.data?.qr || result.result?.qrcode || result.result?.qr;
      
      if (qrValue) {
        return `
          <div style="text-align: center;">
            <p>Scan QR code dengan WhatsApp Anda:</p>
            <img src="${qrValue}" width="200" height="200"><br>
            <p style="font-size: 12px;">QR akan kedaluwarsa dalam 60 detik</p>
          </div>
        `;
      } else {
        // Tampilkan lebih banyak informasi tentang respons
        return `QR berhasil dibuat, tetapi tidak ditemukan URL QR dalam respons. Mohon periksa log eksekusi script atau hubungi penyedia API.<br><br>
               <details>
                 <summary>Detail Respons API (klik untuk membuka)</summary>
                 <pre>${JSON.stringify(result, null, 2)}</pre>
               </details>`;
      }
    } else {
      return 'Gagal membuat QR: ' + result.message;
    }
  } catch (error) {
    return 'Error: ' + error.toString();
  }
}

/**
 * Menampilkan dialog logout device
 */
function showLogoutDeviceDialog() {
  const htmlOutput = HtmlService
    .createHtmlOutput(`
      <form>
        <div style="margin-bottom: 10px;">
          <label for="sender">Nomor Device yang akan dilogout:</label><br>
          <input type="text" id="sender" name="sender" placeholder="628xxxxxxxxxx" style="width: 100%;">
        </div>
        <div style="text-align: center; margin-top: 20px;">
          <input type="button" value="Logout" onclick="logoutDevice()" style="padding: 5px 15px;">
        </div>
      </form>
      <div id="result" style="margin-top: 10px; text-align: center;"></div>
      <script>
        function logoutDevice() {
          const sender = document.getElementById('sender').value;
          
          if (!sender) {
            document.getElementById('result').innerHTML = 'Nomor device harus diisi!';
            return;
          }
          
          if (!confirm('Apakah Anda yakin ingin logout device ini?')) {
            return;
          }
          
          document.getElementById('result').innerHTML = 'Melakukan logout...';
          google.script.run
            .withSuccessHandler(showResult)
            .withFailureHandler(showError)
            .logoutDevice(sender);
        }
        
        function showResult(result) {
          document.getElementById('result').innerHTML = result;
        }
        
        function showError(error) {
          document.getElementById('result').innerHTML = 'Error: ' + error.message;
        }
      </script>
    `)
    .setWidth(400)
    .setHeight(200)
    .setTitle('Logout Device WhatsApp');
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Logout Device WhatsApp');
}

/**
 * Fungsi untuk logout device
 */
function logoutDevice(sender) {
  const settings = getSettings();
  
  const payload = {
    api_key: settings.apiKey,
    sender: sender
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  };
  
  try {
    const response = UrlFetchApp.fetch(`${BASE_URL}/logout-device`, options);
    const result = JSON.parse(response.getContentText());
    
    if (result.status) {
      return 'Device berhasil dilogout!';
    } else {
      return 'Gagal logout device: ' + result.message;
    }
  } catch (error) {
    return 'Error: ' + error.toString();
  }
}

/**
 * Menampilkan dialog untuk kirim pesan massal dari sheet
 */
function showBulkMessageDialog() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Membuat opsi untuk kolom nomor dan pesan
  let numberOptions = '';
  let messageOptions = '';
  
  headers.forEach((header, index) => {
    numberOptions += `<option value="${index+1}">${header}</option>`;
    messageOptions += `<option value="${index+1}">${header}</option>`;
  });
  
  const htmlOutput = HtmlService
    .createHtmlOutput(`
      <form>
        <p style="color: #666;">Pilih kolom yang berisi nomor tujuan dan pesan yang akan dikirim</p>
        <div style="margin-bottom: 10px;">
          <label for="numberColumn">Kolom Nomor Tujuan:</label><br>
          <select id="numberColumn" name="numberColumn" style="width: 100%;">
            ${numberOptions}
          </select>
        </div>
        <div style="margin-bottom: 10px;">
          <label for="messageColumn">Kolom Pesan:</label><br>
          <select id="messageColumn" name="messageColumn" style="width: 100%;">
            ${messageOptions}
          </select>
        </div>
        <div style="margin-bottom: 10px;">
          <label for="startRow">Baris Mulai:</label><br>
          <input type="number" id="startRow" name="startRow" value="2" min="2" style="width: 100%;">
        </div>
        <div style="margin-bottom: 10px;">
          <label for="endRow">Baris Akhir (kosongkan untuk semua):</label><br>
          <input type="number" id="endRow" name="endRow" min="2" style="width: 100%;">
        </div>
        <div style="text-align: center; margin-top: 20px;">
          <input type="button" value="Kirim Pesan Massal" onclick="sendBulkMessages()" style="padding: 5px 15px;">
        </div>
      </form>
      <div id="result" style="margin-top: 10px; text-align: center;"></div>
      <script>
        function sendBulkMessages() {
          const numberColumn = document.getElementById('numberColumn').value;
          const messageColumn = document.getElementById('messageColumn').value;
          const startRow = document.getElementById('startRow').value;
          const endRow = document.getElementById('endRow').value || '';
          
          if (!numberColumn || !messageColumn || !startRow) {
            document.getElementById('result').innerHTML = 'Semua field harus diisi kecuali baris akhir!';
            return;
          }
          
          if (!confirm('Apakah Anda yakin ingin mengirim pesan massal?')) {
            return;
          }
          
          document.getElementById('result').innerHTML = 'Mengirim pesan massal...';
          google.script.run
            .withSuccessHandler(showResult)
            .withFailureHandler(showError)
            .sendBulkMessages(numberColumn, messageColumn, startRow, endRow);
        }
        
        function showResult(result) {
          document.getElementById('result').innerHTML = result;
        }
        
        function showError(error) {
          document.getElementById('result').innerHTML = 'Error: ' + error.message;
        }
      </script>
    `)
    .setWidth(400)
    .setHeight(400)
    .setTitle('Kirim Pesan Massal WhatsApp');
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Kirim Pesan Massal WhatsApp');
}

/**
 * Fungsi untuk mengirim pesan massal dari sheet
 */
function sendBulkMessages(numberColumn, messageColumn, startRow, endRow) {
  const settings = getSettings();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Konversi ke angka
  numberColumn = parseInt(numberColumn);
  messageColumn = parseInt(messageColumn);
  startRow = parseInt(startRow);
  
  // Tentukan baris akhir
  const lastRow = sheet.getLastRow();
  endRow = endRow ? parseInt(endRow) : lastRow;
  
  if (endRow > lastRow) {
    endRow = lastRow;
  }
  
  let successCount = 0;
  let failCount = 0;
  let errorMessages = [];
  
  // Tambahkan kolom status jika belum ada
  let statusColumn = sheet.getLastColumn() + 1;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  if (headers.indexOf('Status Pengiriman') >= 0) {
    statusColumn = headers.indexOf('Status Pengiriman') + 1;
  } else {
    sheet.getRange(1, statusColumn).setValue('Status Pengiriman');
  }
  
  // Tambahkan kolom timestamp jika belum ada
  let timestampColumn = statusColumn + 1;
  
  if (headers.indexOf('Waktu Pengiriman') >= 0) {
    timestampColumn = headers.indexOf('Waktu Pengiriman') + 1;
  } else {
    sheet.getRange(1, timestampColumn).setValue('Waktu Pengiriman');
  }
  
  // Loop melalui baris yang dipilih
  for (let row = startRow; row <= endRow; row++) {
    const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
    const number = rowData[numberColumn - 1];
    const message = rowData[messageColumn - 1];
    
    // Lewati baris jika nomor atau pesan kosong
    if (!number || !message) {
      sheet.getRange(row, statusColumn).setValue('DILEWATI - Data tidak lengkap');
      sheet.getRange(row, timestampColumn).setValue(new Date());
      continue;
    }
    
    try {
      const payload = {
        api_key: settings.apiKey,
        sender: settings.sender,
        number: number.toString().trim(),
        message: message
      };
      
      const options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };
      
      const response = UrlFetchApp.fetch(`${BASE_URL}/send-message`, options);
      const result = JSON.parse(response.getContentText());
      
      // Update status di sheet
      if (result.status) {
        sheet.getRange(row, statusColumn).setValue('SUKSES');
        successCount++;
      } else {
        sheet.getRange(row, statusColumn).setValue('GAGAL - ' + result.message);
        failCount++;
        errorMessages.push(`Baris ${row}: ${result.message}`);
      }
      
      // Update timestamp
      sheet.getRange(row, timestampColumn).setValue(new Date());
      
      // Delay untuk menghindari rate limit (1 detik)
      Utilities.sleep(1000);
    } catch (error) {
      sheet.getRange(row, statusColumn).setValue('ERROR - ' + error.toString());
      sheet.getRange(row, timestampColumn).setValue(new Date());
      failCount++;
      errorMessages.push(`Baris ${row}: ${error.toString()}`);
    }
  }
  
  return `Pengiriman selesai. Sukses: ${successCount}, Gagal: ${failCount}`;
}
