// ============================
// Firebase Configuration
// ============================
var FC = {
  apiKey: "AIzaSyBsgnrbFI4-F_Vy-cD6EB7LRwdJxuAxzL8",
  authDomain: "project-3074788574588051418.firebaseapp.com",
  projectId: "project-3074788574588051418",
  storageBucket: "project-3074788574588051418.firebasestorage.app",
  messagingSenderId: "967388882827",
  appId: "1:967388882827:web:78fc09e700a6c903a5718e"
};

// ============================
// Global State
// ============================
var db = null, cls = [], cu = null, cfl = 'all', uns = null, unsPm = null;
var isDk = true, cfCb = null, dtFlA = false, dtFlFld = 'created';
var hideTm = false, srvOff = 0, autoInt = null, clkInt = null;

// ============================
// Constants
// ============================
var MA = ['يناير','فبراير','مارس','أبريل','مايو','يونيو','يوليو','أغسطس','سبتمبر','أكتوبر','نوفمبر','ديسمبر'];
var STL = { uploaded: 'تم الرفع', pending: 'قيد المراجعة', refunded: 'تم الاسترداد', rejected: 'مرفوض' };
var STC = { uploaded: 'sb-u', pending: 'sb-p', refunded: 'sb-d', rejected: 'sb-r' };
var STI = { uploaded: 'fa-solid fa-cloud-arrow-up', pending: 'fa-solid fa-clock', refunded: 'fa-solid fa-circle-check', rejected: 'fa-solid fa-circle-xmark' };
