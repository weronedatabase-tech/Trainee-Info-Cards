// ==========================================
// config.js - Application Configuration
// ==========================================
// ENVIRONMENT TOGGLE
// Options: 'Exp' (Experimental) | 'Dev' (Development) | 'Prod' (Production)
const ENV = 'Prod';

// ENVIRONMENT API ENDPOINTS (Google Apps Script Web App URLs)
const EXP_URL = 'https://script.google.com/macros/s/AKfycbz5lEH75nTIX2epTYjirnSesn0Lt7HAUrczaEYTL031haBWzNaQSw9FmOowU_5jD1TY/exec';
const DEV_URL = 'https://script.google.com/macros/s/AKfycbz5lEH75nTIX2epTYjirnSesn0Lt7HAUrczaEYTL031haBWzNaQSw9FmOowU_5jD1TY/exec';
const PROD_URL = 'https://script.google.com/macros/s/AKfycbyEop3q0sfBjN4PXt3FoRJL2byI9wpez-Bp0N4PQ8kGJjb3AJqhZoSxyjcmKWlgfZkZIw/exec';
const API_URL = ENV === 'Exp' ? EXP_URL : (ENV === 'Dev' ? DEV_URL : PROD_URL);

// ENVIRONMENT Google Drive Folders (Google Drive Folder IDs) for the trainee photos
const EXP_Drive_Folder_ID = '1nMFek_9bTttYPVW_vlV1eOfawDz3RGy-';
const DEV_Drive_Folder_ID= '1nMFek_9bTttYPVW_vlV1eOfawDz3RGy-';
const PROD_Drive_Folder_ID = '1nMFek_9bTttYPVW_vlV1eOfawDz3RGy-';
const Drive_Folder_ID = ENV === 'Exp' ? EXP_Drive_Folder_ID : (ENV === 'Dev' ? DEV_Drive_Folder_ID : PROD_Drive_Folder_ID);
