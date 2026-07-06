// 알림 자동 닫기
document.querySelectorAll('.alert').forEach(function(el) {
  setTimeout(function() {
    const bs = bootstrap.Alert.getOrCreateInstance(el);
    if (bs) bs.close();
  }, 4000);
});

// 오늘 날짜를 date input 기본값으로
document.querySelectorAll('input[name="completed_at"]').forEach(function(el) {
  if (!el.value) {
    const today = new Date().toISOString().split('T')[0];
    el.value = today;
  }
});
