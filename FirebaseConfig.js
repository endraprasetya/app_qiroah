// Ganti nilai di bawah ini sesuai dengan isi file JSON dari Firebase Anda
const FIREBASE_CONFIG = {
  projectId: "app-qiroahalmadinah", // Contoh: "qiroatul-quran-sd-12345"
  clientEmail: "firebase-adminsdk-fbsvc@app-qiroahalmadinah.iam.gserviceaccount.com", // Contoh: "firebase-adminsdk-xxx@qiroatul-quran-sd-12345.iam.gserviceaccount.com"
  privateKey: "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQCa5ENSvTxR2FvH\n/ptsxDQKT+ODTxrpBddQzfRjUvYH8xE5uxGYtwnXqxmMJcevK6tXCb4KmwVicdaq\nFdgESjOSbgp0fclhfHMlbGMfs1sfufN8ZTrl6IJPyeON6ngaeicWW8eNCsn6bD7m\nteU+/DrXAbVCbDorCJB6bPFSiCZpNYpfjCfniw9P3bfOi7n7mNzteWmW6SFXDzR6\n4vux/rxzU+iLYXqysS4gQks3DQA6mdaWSK125wfBGJixC+bxv+SRaw+Tl4xS91uV\nZu1xrfH1kBQWGTsVSNxc+7caiMmbjL+I/I48CsGmSq7TeGHOrjGcoiViRpCb2Mne\nfOiMdHljAgMBAAECggEAFBHeBqizhK1isxwyr8S6iNgyKYUEqSCodiAW+bm5scJV\nPzmV6EerVFEncHPtbFgIYwZFpCJ9xPRC4EC6qqT3YBK3HcILZ9Ype2C7H/6W9kTw\nykkO+SmynOZoKwQwch8q1RL+XTBX3IV4dF9+/WAkZaX4gnX7PDE8u0hGJVoJPsTN\nc8c6xEtx3ty3dsobGgZgnPJtfhNdqD42Jk2+sgnTU9UXdypHLk++p0XKcu1xkbS/\nqK/JS4JmHX/+qHV2uFWa3H4oP7AMqH+9mlpv4tl1g0eHIAO8XxxxZTgiNe5EYnX2\n+ZsELP5wAFVePo7tMg4TWBEp/sNlc1EP5Q4qQhsnMQKBgQDLlmy59hQCHzpADTpe\nsMtHr54k1fArwMIa6s25vT7qA7KOMCATFcKjx9DbJJQozjBAGGByE48ONfCae8rC\nICoh5Ip/Y4rYCmCxl6BIEhljR2RLY7ENIuI3bKNXG/D82hsscRWibac3lufBdZmK\nPi1hBGToNzkMawMnd5yjFVfucQKBgQDCxH/1Io1rnCdf37LNxR+iA5Y/7iQz4HSr\npyDWyAVqG+P1O2IX5IMb5yDzH3k2ocwr3X5gs130gxvZZTgVmxM9ZOliRih/wGoQ\nZsul7CS/aktyIS9TjzQ/vzIAHcUvo8OKRfs8idy2uBsbB+1pCFAOQ+yGR+qa7Mh+\nf6zENqu3EwKBgQC4UxFNrQaghTVwGInogZkK5m4AW9J/Sp5zUHPC3xy/M2nOYrQn\nk2QOdTMbthqOd2AoJ+nqXtr+useEluAUsppy5UCf5bgnsf2pzOZJWlmyEUk1GKWM\nhzDZjXq1u46M38wKiLhVv4HOSsq51YeWKoVbeWcC+DymiZzpdtpSrVhXgQKBgH3C\nvdC+L8y1NqNsYRsG1qcZY9IB5fRtUUhpEQY+f2iFe8skPfGphpLkydzSjMUe8iO8\n0FaGWvuAEOJtvioNSP4Sg7QGEy5yQso2ZfbMJOpIK1EVZmaxCA7FWPbEQkfBmYhW\nKaRsbEy7Uh1BICX6zR5482e9VhoVHFy5+gd0OaprAoGADwZ1QenIf8bWbwwXhk13\noceZC4Sf/VVNUwcmjYqO4r9tMlm9z5V7lj2K1kQm+CyXf8TA8WDY5J70NA3qDtT7\n3Wyuo+JBQNysWMIS0yR1BAMindmd9a9++f2imDNNXSwJkr+dAq8VKDDiKMJtFkVX\nQSXS1pN10EjTz0mCOErmMuQ=\n-----END PRIVATE KEY-----\n" 
  // Pastikan private key diawali "-----BEGIN PRIVATE KEY-----\n" 
  // dan diakhiri "\n-----END PRIVATE KEY-----\n"
  // Gunakan format string bersambung seperti di project sebelumnya jika perlu
};

// Fungsi ini akan dipanggil oleh file lain untuk berkomunikasi dengan database
function getFirestoreDB() {
  return FirestoreApp.getFirestore(
    FIREBASE_CONFIG.clientEmail,
    FIREBASE_CONFIG.privateKey,
    FIREBASE_CONFIG.projectId
  );
}