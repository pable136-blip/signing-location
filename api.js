// Apps Script 웹 앱 주소 ↔ 배포 ID.
// 부스 화면은 QR 에 짧은 배포 ID 만 싣고, 출석 화면은 그 ID 로 항상 script.google.com 주소를 만든다
// (QR 을 조작해 엉뚱한 서버로 개인정보를 보내게 하는 것을 막기 위해 다른 도메인은 받지 않는다).
window.AttendanceApi = {
  idFromUrl: function (u) {
    var m = String(u || '').trim().match(
      /^https:\/\/script\.google\.com\/(?:a\/macros\/[^\/]+|macros)\/s\/([A-Za-z0-9_-]{20,})\/exec\/?$/);
    return m ? m[1] : '';
  },
  urlFromId: function (id) {
    return /^[A-Za-z0-9_-]{20,}$/.test(id || '') ? 'https://script.google.com/macros/s/' + id + '/exec' : '';
  }
};
