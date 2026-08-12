const ENV_CONFIG = {
  ACTIVE_ENV: 'DEV', // Set to 'PROD', 'DEV', or 'EXP'
  
  URLS: {
    DEV: 'https://script.google.com/macros/s/AKfycby48gbzI_4V0TEJ0Gra4Qb_J3xywBA6A792d2reGx0QWUx-6QFEKRWBTmr8mGG86osg/exec',
    PROD: 'https://script.google.com/macros/s/AKfycbz4OLZtR2lX97MrGZVaNg13Lrzvwgy7mBfQr7PgoQGK617sL8ZCkKvZD2hIZodus-O_/exec',
    EXP: 'https://script.google.com/macros/s/AKfycby48gbzI_4V0TEJ0Gra4Qb_J3xywBA6A792d2reGx0QWUx-6QFEKRWBTmr8mGG86osg/exec' // Currently using DEV URL as placeholder
  },
  
  SHEETS: {
    DEV: '1E1GPV36RLHn7p4gHmdHB2zedtZJF2_Zb3ncGXZZTy5Y',
    PROD: '1IbxJY59urIChYaLrwURWrOEBBrLNDiOqDO-w6LwP1xI',
    EXP: '1E1GPV36RLHn7p4gHmdHB2zedtZJF2_Zb3ncGXZZTy5Y'
  },
  
  DRIVE_FOLDER_ID: '1nMFek_9bTttYPVW_vlV1eOfawDz3RGy-'
};

// Export configuration so it works safely in both Google Apps Script and modern Node/Browser contexts.
if (typeof module !== 'undefined' && module.exports) {
    module.exports = ENV_CONFIG;
}
