/* ---------------- CONFIGURATION ---------------- */
const SHEET_ID = "1qT2CIC-z5UDYtDbw12swRlgR8VfQBHZQCJrMPXLAc-0"; // *** PASTE YOUR GOOGLE SHEET ID HERE ***

/* ---------------- SERVE HTML ---------------- */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Office Attendance')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------------- API HANDLERS ---------------- */

// 1. LOGIN
function loginUser(email, password) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == email && data[i][1] == password) {
      return { 
        success: true, 
        name: data[i][2], 
        role: data[i][3], 
        email: data[i][0] 
      };
    }
  }
  return { success: false, message: "Invalid Credentials" };
}

// 2. GENERATE ROTATING QR TOKEN (ADMIN SIDE)
function generateToken() {
  const token = Math.random().toString(36).substring(2, 15);
  // Store token in Script Properties with timestamp
  const props = PropertiesService.getScriptProperties();
  props.setProperty('CURRENT_TOKEN', token);
  props.setProperty('TOKEN_TIME', new Date().getTime());
  return token;
}

// 3. MARK ATTENDANCE (EMPLOYEE SIDE)
function markAttendance(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const settings = ss.getSheetByName('Settings').getDataRange().getValues();
  const props = PropertiesService.getScriptProperties();
  
  // A. Validate Token (Time limit 45s to allow scan delay)
  const serverToken = props.getProperty('CURRENT_TOKEN');
  const tokenTime = parseInt(props.getProperty('TOKEN_TIME'));
  const now = new Date().getTime();
  
  if (data.qrToken !== serverToken || (now - tokenTime > 45000)) {
    return { success: false, message: "QR Code Expired. Scan again." };
  }

  // B. Validate IP (Wifi Check)
  const requiredIP = settings.find(r => r[0] === 'OFFICE_IP')[1];
  // Note: Standard IP check is client-provided in this lightweight MVP. 
  // For strict security, backend header inspection requires detailed deployment.
  if (data.userIP.trim() !== requiredIP.toString().trim()) {
    return { success: false, message: "Please connect to Office WiFi." };
  }

  // C. Validate Location (Geofence)
  const officeLat = settings.find(r => r[0] === 'OFFICE_LAT')[1];
  const officeLng = settings.find(r => r[0] === 'OFFICE_LNG')[1];
  const allowedRadius = settings.find(r => r[0] === 'ALLOWED_RADIUS')[1];
  
  const distance = getDistanceFromLatLonInM(data.lat, data.lng, officeLat, officeLng);
  
  if (distance > allowedRadius) {
    return { success: false, message: `You are ${Math.round(distance)}m away. Move closer.` };
  }

  // D. Log to Sheet
  ss.getSheetByName('Attendance').appendRow([
    new Date(), 
    data.name, 
    data.email, 
    "PRESENT", 
    `IP: ${data.userIP} | Dist: ${Math.round(distance)}m`
  ]);

  return { success: true, message: "Attendance Marked Successfully!" };
}

// HELPER: Haversine Formula for Distance
function getDistanceFromLatLonInM(lat1, lon1, lat2, lon2) {
  var R = 6371000; // Radius of earth in meters
  var dLat = deg2rad(lat2-lat1);
  var dLon = deg2rad(lon2-lon1); 
  var a = 
    Math.sin(dLat/2) * Math.sin(dLat/2) +
    Math.cos(deg2rad(lat1)) * Math.cos(deg2rad(lat2)) * Math.sin(dLon/2) * Math.sin(dLon/2); 
  var c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a)); 
  var d = R * c; 
  return d;
}

function deg2rad(deg) {
  return deg * (Math.PI/180);
}
